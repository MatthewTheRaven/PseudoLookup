/**
 * MultiTableLookup component - Main container for the lookup control
 * Displays badges for selected records and provides lookup functionality
 */

import * as React from 'react';
import { ILookupRecord, ITableIconCache } from '../types';
import { LookupBadge } from './LookupBadge';

export interface IMultiTableLookupProps {
/** Available table logical names for lookup */
lookupTables: string[];
/** Whether multi-select is allowed */
allowMultiSelect: boolean;
/** Whether clicking badges navigates to the record */
enableLinking: boolean;
/** Whether the control is disabled/read-only */
isDisabled: boolean;
/** Whether maximum selections has been reached */
isAtMaxSelections: boolean;
/** Placeholder text shown on hover when empty */
placeholderText: string;
    /** State version from PCF control to trigger re-renders */
    stateVersion: number;
    /** Getter for currently selected records (always returns latest) */
    getSelectedRecords: () => ILookupRecord[];
    /** Getter for error message (always returns latest) */
    getErrorMessage: () => string | undefined;
    /** Getter for loading state (always returns latest) */
    getIsLoading: () => boolean;
    /** Getter for table icon cache (always returns latest) */
    getTableIconCache: () => ITableIconCache;
    /** Callback to register state updater function */
    onStateUpdaterReady: (updater: (version: number) => void) => void;
    /** Callback when user clicks the lookup button */
    onLookupClick: () => void;
    /** Callback when user removes a record */
    onRemoveRecord: (recordId: string) => void;
    /** Callback when user clicks a badge to navigate */
    onNavigateToRecord: (record: ILookupRecord) => void;
}

/**
 * MultiTableLookup renders the lookup control with badges and lookup button
 */
export const MultiTableLookup: React.FC<IMultiTableLookupProps> = ({
    lookupTables,
    enableLinking,
    isDisabled,
    isAtMaxSelections,
    placeholderText,
    stateVersion,
    getSelectedRecords,
    getErrorMessage,
    getIsLoading,
    getTableIconCache,
    onStateUpdaterReady,
    onLookupClick,
    onRemoveRecord,
    onNavigateToRecord
}) => {
    // State for hover detection on empty control
    const [isHovered, setIsHovered] = React.useState(false);
    
    // Internal state version to trigger re-renders when PCF control state changes
    const [, setInternalStateVersion] = React.useState(stateVersion);
    
    // Register the state updater callback with the PCF control on mount
    React.useEffect(() => {
        onStateUpdaterReady((version: number) => {
            setInternalStateVersion(version);
        });
    }, [onStateUpdaterReady]);
    
    // Get the latest state from the PCF control via getters
    const selectedRecords = getSelectedRecords();
    const errorMessage = getErrorMessage();
    const isLoading = getIsLoading();
    const tableIconCache = getTableIconCache();
    
    /**
     * Handles mouse enter on the container
     */
    const handleMouseEnter = React.useCallback(() => {
        setIsHovered(true);
    }, []);
    
    /**
     * Handles mouse leave on the container
     */
    const handleMouseLeave = React.useCallback(() => {
        setIsHovered(false);
    }, []);

    /**
     * Handles click on the lookup button or container
     */
    const handleLookupClick = React.useCallback(() => {
        if (!isDisabled && !isAtMaxSelections) {
            onLookupClick();
        }
    }, [isDisabled, isAtMaxSelections, onLookupClick]);

    /**
     * Handles click on the container area (not on badges)
     */
    const handleContainerClick = React.useCallback((event: React.MouseEvent<HTMLDivElement>) => {
        // Only trigger if clicking directly on the container or content area, not on badges
        const target = event.target as HTMLElement;
        const isClickOnBadge = target.closest('.lookup-badge');
        if (!isClickOnBadge && !isDisabled && !isAtMaxSelections) {
            onLookupClick();
        }
    }, [isDisabled, isAtMaxSelections, onLookupClick]);

    /**
     * Handles keyboard interaction on the container
     */
    const handleContainerKeyDown = React.useCallback((event: React.KeyboardEvent<HTMLDivElement>) => {
        // Only trigger if the container itself is focused (not child elements)
        if (!isDisabled && !isAtMaxSelections && (event.key === 'Enter' || event.key === ' ') && event.target === event.currentTarget) {
            event.preventDefault();
            onLookupClick();
        }
    }, [isDisabled, isAtMaxSelections, onLookupClick]);

    // Determine container classes
    const containerClasses = ['pseudo-lookup-container'];
    if (isDisabled) {
        containerClasses.push('disabled');
    }
    if (errorMessage) {
        containerClasses.push('error');
    }

    // Determine if we should show the empty state placeholder
    const showPlaceholder = !isLoading && selectedRecords.length === 0;
    
    // Determine placeholder text based on hover state
    const displayedPlaceholderText = isHovered ? placeholderText : '---';

    return (
        <div className="pseudo-lookup-wrapper">
            <div 
                className={containerClasses.join(' ')}
                role="combobox"
                aria-expanded="false"
                aria-haspopup="dialog"
                aria-label={`Lookup field with ${selectedRecords.length} selected record${selectedRecords.length !== 1 ? 's' : ''}`}
                aria-disabled={isDisabled}
                tabIndex={isDisabled ? -1 : 0}
                onClick={handleContainerClick}
                onKeyDown={handleContainerKeyDown}
                onMouseEnter={handleMouseEnter}
                onMouseLeave={handleMouseLeave}
                style={{ cursor: isDisabled ? 'not-allowed' : (isAtMaxSelections ? 'default' : 'pointer') }}
            >
                {/* Content area - badges or placeholder */}
                <div 
                    className="pseudo-lookup-content"
                    role="list"
                    aria-label="Selected records"
                >
                    {showPlaceholder ? (
                        // Empty state placeholder with hover behavior
                        <span className={`pseudo-lookup-placeholder ${isHovered ? 'hovered' : ''}`}>
                            {displayedPlaceholderText}
                        </span>
                    ) : (
                        // Render badges for each selected record
                        selectedRecords.map((record) => (
                            <LookupBadge
                                key={record.id}
                                record={record}
                                //iconUrl={record.iconUrl ?? tableIconCache[record.table.toLowerCase()] ?? undefined}
                                enableLinking={enableLinking}
                                isDisabled={isDisabled}
                                isLoading={record.isLoading}
                                onRemove={onRemoveRecord}
                                onNavigate={onNavigateToRecord}
                            />
                        ))
                    )}
                </div>

                {/* Lookup icon - magnifying glass (hidden when disabled or at max selections) */}
                {!isDisabled && !isAtMaxSelections && (
                    <span
                        className="pseudo-lookup-icon symbolFont SearchButton-symbol"
                        aria-hidden="true"
                    />
                )}
            </div>

            {/* Error message display */}
            {errorMessage && (
                <div className="pseudo-lookup-error" role="alert">
                    <svg 
                        className="pseudo-lookup-error-icon" 
                        viewBox="0 0 12 12" 
                        xmlns="http://www.w3.org/2000/svg"
                        aria-hidden="true"
                    >
                        <path d="M6 0a6 6 0 1 0 0 12A6 6 0 0 0 6 0zm0 10.5a.75.75 0 1 1 0-1.5.75.75 0 0 1 0 1.5zm.75-3a.75.75 0 0 1-1.5 0V3a.75.75 0 0 1 1.5 0v4.5z" />
                    </svg>
                    <span>{errorMessage}</span>
                </div>
            )}
        </div>
    );
};

export default MultiTableLookup;
