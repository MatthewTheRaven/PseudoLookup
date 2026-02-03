/**
 * PseudoLookup - Multi-Table Lookup PCF Control
 * 
 * A PowerApps Component Framework control that binds to a SingleLine.Text field
 * and provides a multi-table lookup experience similar to native Dataverse lookups.
 * 
 * Data is stored as semicolon-delimited string: {guid}:{tablename};{guid}:{tablename}
 */

import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { MultiTableLookup } from "./components/MultiTableLookup";
import { ILookupRecord, ITableIconCache } from "./types";
import {
    parseBoundValue,
    serializeToBoundValue,
    parseLookupTables,
    generateAllExclusionFilters,
    removeRecordById,
    normalizeGuid
} from "./utils/dataUtils";
import {
    retrieveAllRecords,
    openLookupDialog,
    openRecordForm,
    getTableIconsForTables
} from "./utils/webApiUtils";
import * as React from "react";

/**
 * Main control class implementing the PCF ReactControl interface
 */
export class PseudoLookup implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    
    // Internal state
    private selectedRecords: ILookupRecord[] = [];
    private errorMessage: string | undefined;
    private boundFieldValue = '';
    private recordNamesValue = '';
    private isInitialized = false;
    private lastBoundValue: string | null = null;
    private isLoadingInProgress = false;
    
    // Table icon cache
    private tableIconCache: ITableIconCache = {};
    private iconCacheLoaded = false;
    
    // Debounce flag to prevent multiple rapid navigations
    private isNavigating = false;
    
    // State version counter and updater for React re-renders
    private stateVersion = 0;
    private stateUpdater: ((version: number) => void) | null = null;

    /**
     * Empty constructor.
     */
    constructor() {
        // Empty - initialization happens in init()
    }

    /**
     * Used to initialize the control instance.
     * Kicks off Web API calls to retrieve record names if bound field has values.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Triggers a re-render by incrementing the state version and calling the updater
     */
    private triggerRerender(): void {
        this.stateVersion++;
        if (this.stateUpdater) {
            this.stateUpdater(this.stateVersion);
        }
    }

    /**
     * Sets the state updater callback from the React component
     */
    private setStateUpdater(updater: (version: number) => void): void {
        this.stateUpdater = updater;
    }

    /**
     * Returns the current selected records (used by React component to get latest state)
     */
    private getSelectedRecords(): ILookupRecord[] {
        return this.selectedRecords;
    }

    /**
     * Returns the current error message (used by React component to get latest state)
     */
    private getErrorMessage(): string | undefined {
        return this.errorMessage;
    }

    /**
     * Returns the current loading state (used by React component to get latest state)
     */
    private getIsLoading(): boolean {
        return this.isLoadingInProgress;
    }

    /**
     * Returns the current table icon cache (used by React component to get latest state)
     */
    private getTableIconCache(): ITableIconCache {
        return this.tableIconCache;
    }

    /**
     * Loads table icons for configured lookup tables
     */
    private async loadTableIcons(context: ComponentFramework.Context<IInputs>): Promise<void> {
        if (this.iconCacheLoaded) {
            return;
        }
        
        const lookupTablesStr = context.parameters.lookupTables?.raw ?? '';
        const lookupTables = parseLookupTables(lookupTablesStr);
        
        if (lookupTables.length === 0) {
            return;
        }
        
        try {
            this.tableIconCache = await getTableIconsForTables(context, lookupTables);
            this.iconCacheLoaded = true;
            this.triggerRerender();
        } catch (error) {
            console.warn('[PseudoLookup] Error loading table icons:', error);
            this.iconCacheLoaded = true;
        }
    }

    /**
     * Loads initial data from the bound field value
     * Immediately creates placeholder badges, then retrieves record names via Web API
     * Handles external value changes that occur during loading
     */
    private async loadInitialData(context: ComponentFramework.Context<IInputs>): Promise<void> {
        const boundValue = context.parameters.boundField?.raw ?? '';
        
        // Skip if already loading the same value
        if (this.isLoadingInProgress) {
            return;
        }
        
        // Handle empty bound field
        if (!boundValue || boundValue.trim() === '') {
            this.selectedRecords = [];
            this.recordNamesValue = '';
            this.boundFieldValue = '';
            this.isInitialized = true;
            this.lastBoundValue = '';
            this.triggerRerender();
            return;
        }
        
        // Skip if value hasn't changed since last load
        if (boundValue === this.lastBoundValue && this.isInitialized) {
            return;
        }
        
        // Store the value we're loading so we can detect if it changes during load
        const loadingValue = boundValue;
        this.lastBoundValue = boundValue;
        this.isLoadingInProgress = true;
        this.errorMessage = undefined;
        
        try {
            // Parse the bound value to get id:table pairs
            const parsedRecords = parseBoundValue(boundValue);
            
            if (parsedRecords.length === 0) {
                this.selectedRecords = [];
                this.recordNamesValue = '';
                this.boundFieldValue = '';
                this.isLoadingInProgress = false;
                this.isInitialized = true;
                this.triggerRerender();
                return;
            }
            
            // Immediately create placeholder badges with loading state
            this.selectedRecords = parsedRecords.map(record => ({
                id: record.id,
                table: record.table,
                name: 'Loading...',
                isLoading: true
            }));
            this.isInitialized = true;
            this.triggerRerender();
            
            // Retrieve all records via Web API
            const result = await retrieveAllRecords(context, parsedRecords);
            
            // Check if the bound value changed while we were loading
            const currentBoundValue = this.context.parameters.boundField?.raw ?? '';
            if (currentBoundValue !== loadingValue) {
                this.isLoadingInProgress = false;
                void this.loadInitialData(this.context);
                return;
            }
            
            // Update records with actual data
            this.selectedRecords = result.validRecords;
            this.recordNamesValue = this.serializeRecordNames();
            
            // If any records were deleted, update the bound field
            if (result.hasDeletedRecords) {
                this.boundFieldValue = serializeToBoundValue(this.selectedRecords);
                this.lastBoundValue = this.boundFieldValue;
                this.notifyOutputChanged();
            }
            
        } catch (error) {
            console.error('[PseudoLookup] Error loading records:', error);
            this.errorMessage = 'Error loading records. Please refresh.';
            this.selectedRecords = [];
            this.recordNamesValue = '';
        } finally {
            this.isLoadingInProgress = false;
            this.triggerRerender();
        }
    }

    /**
     * Called when any value in the property bag has changed.
     * Returns the React element to render.
     * This is called by the framework when the bound field value changes externally
     * (e.g., via getAttribute('fieldname').setValue() or form load)
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        
        // Check if bound field value changed externally or if we need initial load
        const currentBoundValue = context.parameters.boundField?.raw ?? '';
        const valueChanged = currentBoundValue !== this.lastBoundValue;
        
        // Trigger data load if not initialized or value changed externally
        if (!this.isInitialized || (valueChanged && !this.isLoadingInProgress)) {
            void this.loadInitialData(context);
        }
        
        // Load table icons (async, non-blocking)
        if (!this.iconCacheLoaded) {
            void this.loadTableIcons(context);
        }
        
        // Get control properties
        const lookupTablesStr = context.parameters.lookupTables?.raw ?? '';
        const lookupTables = parseLookupTables(lookupTablesStr);
        const allowMultiSelect = context.parameters.allowMultiSelect?.raw === true;
        const enableLinking = context.parameters.enableLinking?.raw === true;
        const placeholderText = context.parameters.placeholderText?.raw ?? 'Look for records';
        const maxSelections = context.parameters.maxSelections?.raw;
        
        // Check if we're at max selections (for disabling lookup)
        const isAtMaxSelections = allowMultiSelect && maxSelections && maxSelections > 0 
            ? this.selectedRecords.length >= maxSelections 
            : false;
        
        // Check if control is disabled
        const isDisabled = context.mode.isControlDisabled || !context.mode.isVisible;
        
        return React.createElement(MultiTableLookup, {
            lookupTables: lookupTables,
            allowMultiSelect: allowMultiSelect,
            enableLinking: enableLinking,
            isDisabled: isDisabled,
            isAtMaxSelections: isAtMaxSelections,
            placeholderText: placeholderText,
            stateVersion: this.stateVersion,
            getSelectedRecords: () => this.getSelectedRecords(),
            getErrorMessage: () => this.getErrorMessage(),
            getIsLoading: () => this.getIsLoading(),
            getTableIconCache: () => this.getTableIconCache(),
            onStateUpdaterReady: (updater: (version: number) => void) => { this.setStateUpdater(updater); },
            onLookupClick: () => { void this.handleLookupClick(); },
            onRemoveRecord: this.handleRemoveRecord.bind(this),
            onNavigateToRecord: (record: ILookupRecord) => { void this.handleNavigateToRecord(record); }
        });
    }

    /**
     * Handles click on the lookup button
     * Opens the lookup dialog with proper filters
     */
    private async handleLookupClick(): Promise<void> {
        if (this.isLoadingInProgress) return;
        
        // Check if we've reached max selections
        const allowMultiSelect = this.context.parameters.allowMultiSelect?.raw === true;
        const maxSelections = this.context.parameters.maxSelections?.raw;
        
        if (allowMultiSelect && maxSelections && maxSelections > 0) {
            if (this.selectedRecords.length >= maxSelections) {
                this.errorMessage = `Maximum of ${maxSelections} selection${maxSelections > 1 ? 's' : ''} reached`;
                this.triggerRerender();
                return;
            }
        }
        
        try {
            // Get lookup configuration
            const lookupTablesStr = this.context.parameters.lookupTables?.raw ?? '';
            const lookupTables = parseLookupTables(lookupTablesStr);
            const disableMru = this.context.parameters.disableMru?.raw === true;
            
            if (lookupTables.length === 0) {
                this.errorMessage = 'No lookup tables configured';
                this.triggerRerender();
                return;
            }
            
            // Determine defaultEntityType - only set if explicitly configured or single table
            let defaultEntityType: string | undefined;
            const configuredDefault = this.context.parameters.defaultEntityType?.raw?.trim().toLowerCase() ?? '';
            
            if (lookupTables.length === 1) {
                // Only one table, use it as default
                defaultEntityType = lookupTables[0];
            } else if (configuredDefault && lookupTables.includes(configuredDefault)) {
                // Multiple tables with valid configured default
                defaultEntityType = configuredDefault;
            }
            // If multiple tables and no valid configured default, leave undefined
            
            // Generate exclusion filters to prevent selecting already-selected records
            const filters = generateAllExclusionFilters(lookupTables, this.selectedRecords);
            
            // Open the lookup dialog
            const results = await openLookupDialog(
                this.context,
                lookupTables,
                allowMultiSelect,
                filters,
                defaultEntityType,
                disableMru
            );
            
            // Handle results
            if (results && results.length > 0) {
                this.processLookupResults(results, allowMultiSelect);
            }
            
        } catch (error) {
            console.error('[PseudoLookup] Error opening lookup dialog:', error);
            // Don't show error for user cancellation
            if (!this.isUserCancelledError(error)) {
                this.errorMessage = 'Error opening lookup dialog';
                this.triggerRerender();
            }
        }
    }

    /**
     * Processes results from the lookup dialog
     */
    private processLookupResults(
        results: ComponentFramework.LookupValue[], 
        allowMultiSelect: boolean
    ): void {
        // Convert lookup results to our record format
        const newRecords: ILookupRecord[] = results.map(result => ({
            id: normalizeGuid(result.id),
            table: result.entityType.toLowerCase(),
            name: result.name ?? 'Unnamed Record',
            iconUrl: this.tableIconCache[result.entityType.toLowerCase()] ?? undefined
        }));
        
        if (allowMultiSelect) {
            // Add to existing selection (filter duplicates just in case)
            const existingIds = new Set(this.selectedRecords.map(r => normalizeGuid(r.id)));
            const uniqueNewRecords = newRecords.filter(r => !existingIds.has(normalizeGuid(r.id)));
            let combinedRecords = [...this.selectedRecords, ...uniqueNewRecords];
            
            // Check if we need to enforce maxSelections limit
            const maxSelections = this.context.parameters.maxSelections?.raw;
            if (maxSelections && maxSelections > 0 && combinedRecords.length > maxSelections) {
                const excessCount = combinedRecords.length - maxSelections;
                combinedRecords = combinedRecords.slice(0, maxSelections);
                this.errorMessage = `Too many selections. Only the first ${maxSelections} record${maxSelections > 1 ? 's were' : ' was'} retained (${excessCount} removed).`;
            }
            
            this.selectedRecords = combinedRecords;
        } else {
            // Replace existing selection
            this.selectedRecords = newRecords.slice(0, 1); // Take only first record
        }
        
        // Update bound field values
        this.boundFieldValue = serializeToBoundValue(this.selectedRecords);
        this.recordNamesValue = this.serializeRecordNames();
        this.lastBoundValue = this.boundFieldValue;

        this.notifyOutputChanged();
        this.triggerRerender();
    }

    /**
     * Handles removal of a record from the selection
     */
    private handleRemoveRecord(recordId: string): void {
        this.selectedRecords = removeRecordById(recordId, this.selectedRecords);
        this.boundFieldValue = serializeToBoundValue(this.selectedRecords);
        this.recordNamesValue = this.serializeRecordNames();
        this.lastBoundValue = this.boundFieldValue;
        this.errorMessage = undefined;
        this.notifyOutputChanged();
        this.triggerRerender();
    }

    /**
     * Handles navigation to a record when badge is clicked
     * Opens in modal or navigates based on openLinksInModal setting
     */
    private async handleNavigateToRecord(record: ILookupRecord): Promise<void> {
        // Debounce to prevent multiple rapid opens
        if (this.isNavigating) return;
        
        this.isNavigating = true;
        
        try {
            const openInModal = this.context.parameters.openLinksInModal?.raw === true;
            await openRecordForm(this.context, record.table, record.id, openInModal);
        } catch (error) {
            // Log error but don't show to user - common for users to cancel navigation
            console.warn('[PseudoLookup] Navigation cancelled or failed:', error);
        } finally {
            // Reset navigation flag after a short delay
            setTimeout(() => {
                this.isNavigating = false;
            }, 500);
        }
    }

    /**
     * Checks if an error indicates user cancellation
     */
    private isUserCancelledError(error: unknown): boolean {
        if (typeof error === 'object' && error !== null) {
            const err = error as { code?: number; message?: string };
            if (err.code === 1 || err.code === -2147020463) {
                return true;
            }
            if (err.message?.toLowerCase().includes('cancel')) {
                return true;
            }
        }
        return false;
    }

    /**
     * Serializes record names to a semicolon-delimited string
     */
    private serializeRecordNames(): string {
        if (!this.selectedRecords || this.selectedRecords.length === 0) {
            return '';
        }
        return this.selectedRecords
            .filter(r => !r.isLoading) // Exclude records still loading
            .map(r => r.name)
            .join(';');
    }

    /**
     * Returns the outputs to be written to the bound fields
     */
    public getOutputs(): IOutputs {
        return {
            boundField: this.boundFieldValue,
            recordNamesOutput: this.recordNamesValue,
            selectionCount: this.selectedRecords.filter(r => !r.isLoading).length
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree.
     * Cleanup resources and event listeners.
     */
    public destroy(): void {
        this.selectedRecords = [];
        this.isLoadingInProgress = false;
        this.errorMessage = undefined;
        this.boundFieldValue = '';
        this.recordNamesValue = '';
        this.isInitialized = false;
        this.lastBoundValue = null;
        this.isNavigating = false;
        this.tableIconCache = {};
        this.iconCacheLoaded = false;
    }
}
