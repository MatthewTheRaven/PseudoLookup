/**
 * Type definitions for the Multi-Table Lookup PCF Control
 */

/**
 * Represents a single selected lookup record
 */
export interface ILookupRecord {
    /** The GUID of the record */
    id: string;
    /** The logical name of the table (entity) */
    table: string;
    /** The primary name field value of the record */
    name: string;
    /** Optional icon URL for the table */
    iconUrl?: string;
    /** Whether this record is still loading its data */
    isLoading?: boolean;
}

/**
 * Represents a parsed record from the bound field (without name)
 */
export interface IParsedRecord {
    /** The GUID of the record */
    id: string;
    /** The logical name of the table (entity) */
    table: string;
}

/**
 * Represents the result from Xrm.Utility.lookupObjects
 */
export interface ILookupResult {
    id: string;
    name: string;
    entityType: string;
}

/**
 * Cache for table icon URLs
 */
export type ITableIconCache = Record<string, string | null>;

/**
 * Props for the main MultiTableLookup component
 */
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
 * Props for the LookupBadge component
 */
export interface ILookupBadgeProps {
    /** The record data */
    record: ILookupRecord;
    /** Whether clicking the badge navigates to the record */
    enableLinking: boolean;
    /** Whether the control is disabled */
    isDisabled: boolean;
    /** Whether this specific record is still loading */
    isLoading?: boolean;
    /** Callback when delete button is clicked */
    onRemove: (recordId: string) => void;
    /** Callback when badge text is clicked (for navigation) */
    onNavigate: (record: ILookupRecord) => void;
}

/**
 * Filter object for lookupObjects
 */
export interface ILookupFilter {
    filterXml: string;
    entityLogicalName: string;
}

/**
 * Options for the lookup dialog
 */
export interface ILookupOptions {
    entityTypes: string[];
    allowMultiSelect: boolean;
    defaultEntityType?: string;
    disableMru?: boolean;
    filters?: ILookupFilter[];
}

/**
 * State for the MultiTableLookup control
 */
export interface IControlState {
    selectedRecords: ILookupRecord[];
    isLoading: boolean;
    errorMessage?: string;
}

/**
 * Result from Web API record retrieval
 */
export interface IRecordRetrievalResult {
    /** Successfully retrieved records */
    validRecords: ILookupRecord[];
    /** Records that failed to retrieve (deleted/inaccessible) */
    deletedRecordIds: string[];
    /** Whether any records were removed */
    hasDeletedRecords: boolean;
}

/**
 * Table metadata for primary name field detection
 */
export interface ITableMetadata {
    logicalName: string;
    primaryNameAttribute: string;
    primaryIdAttribute: string;
}

/**
 * Common table primary name field mappings
 */
export const COMMON_PRIMARY_NAME_FIELDS: Record<string, string> = {
    'contact': 'fullname',
    'account': 'name',
    'lead': 'fullname',
    'opportunity': 'name',
    'systemuser': 'fullname',
    'team': 'name',
    'businessunit': 'name',
    'task': 'subject',
    'phonecall': 'subject',
    'email': 'subject',
    'appointment': 'subject',
    'letter': 'subject',
    'fax': 'subject',
    'activitypointer': 'subject',
    'annotation': 'subject',
    'incident': 'title',
    'knowledgearticle': 'title',
    'queue': 'name',
    'product': 'name',
    'pricelevel': 'name',
    'campaign': 'name',
    'list': 'listname',
    'goal': 'title',
    'transactioncurrency': 'currencyname'
};

/**
 * Represents an icon definition - either a CSS class-based symbol font or an image URL
 */
export interface ITableIconDefinition {
    /** Type of icon: 'symbol' for CSS symbol fonts, 'image' for URL-based icons */
    type: 'symbol' | 'image';
    /** For 'symbol' type: CSS classes to apply. For 'image' type: the image URL */
    value: string;
}

/**
 * Out-of-the-box table icon mappings using Dataverse symbol fonts
 * These tables use CSS classes instead of image URLs for their icons
 * Format: "crmSymbolFont entity-symbol {EntityName} pa-pd"
 */
export const OOB_TABLE_ICON_CLASSES: Record<string, ITableIconDefinition> = {
    'systemuser': { type: 'symbol', value: 'crmSymbolFont entity-symbol Systemuser pa-pd' },
    'account': { type: 'symbol', value: 'crmSymbolFont entity-symbol Account pa-pd' },
    'contact': { type: 'symbol', value: 'crmSymbolFont entity-symbol Contact pa-pd' },
    'lead': { type: 'symbol', value: 'crmSymbolFont entity-symbol Lead pa-pd' },
    'opportunity': { type: 'symbol', value: 'crmSymbolFont entity-symbol Opportunity pa-pd' },
    'team': { type: 'symbol', value: 'crmSymbolFont entity-symbol Team pa-pd' },
    'businessunit': { type: 'symbol', value: 'crmSymbolFont entity-symbol Businessunit pa-pd' },
    'task': { type: 'symbol', value: 'crmSymbolFont entity-symbol Task pa-pd' },
    'phonecall': { type: 'symbol', value: 'crmSymbolFont entity-symbol Phonecall pa-pd' },
    'email': { type: 'symbol', value: 'crmSymbolFont entity-symbol Email pa-pd' },
    'appointment': { type: 'symbol', value: 'crmSymbolFont entity-symbol Appointment pa-pd' },
    'letter': { type: 'symbol', value: 'crmSymbolFont entity-symbol Letter pa-pd' },
    'fax': { type: 'symbol', value: 'crmSymbolFont entity-symbol Fax pa-pd' },
    'activitypointer': { type: 'symbol', value: 'crmSymbolFont entity-symbol Activitypointer pa-pd' },
    'annotation': { type: 'symbol', value: 'crmSymbolFont entity-symbol Annotation pa-pd' },
    'incident': { type: 'symbol', value: 'crmSymbolFont entity-symbol Incident pa-pd' },
    'knowledgearticle': { type: 'symbol', value: 'crmSymbolFont entity-symbol Knowledgearticle pa-pd' },
    'queue': { type: 'symbol', value: 'crmSymbolFont entity-symbol Queue pa-pd' },
    'product': { type: 'symbol', value: 'crmSymbolFont entity-symbol Product pa-pd' },
    'pricelevel': { type: 'symbol', value: 'crmSymbolFont entity-symbol Pricelevel pa-pd' },
    'campaign': { type: 'symbol', value: 'crmSymbolFont entity-symbol Campaign pa-pd' },
    'list': { type: 'symbol', value: 'crmSymbolFont entity-symbol List pa-pd' },
    'goal': { type: 'symbol', value: 'crmSymbolFont entity-symbol Goal pa-pd' },
    'transactioncurrency': { type: 'symbol', value: 'crmSymbolFont entity-symbol Transactioncurrency pa-pd' },
    'invoice': { type: 'symbol', value: 'crmSymbolFont entity-symbol Invoice pa-pd' },
    'quote': { type: 'symbol', value: 'crmSymbolFont entity-symbol Quote pa-pd' },
    'salesorder': { type: 'symbol', value: 'crmSymbolFont entity-symbol Salesorder pa-pd' },
    'contract': { type: 'symbol', value: 'crmSymbolFont entity-symbol Contract pa-pd' },
    'connection': { type: 'symbol', value: 'crmSymbolFont entity-symbol Connection pa-pd' },
    'competitor': { type: 'symbol', value: 'crmSymbolFont entity-symbol Competitor pa-pd' },
    'territory': { type: 'symbol', value: 'crmSymbolFont entity-symbol Territory pa-pd' },
    'recurringappointmentmaster': { type: 'symbol', value: 'crmSymbolFont entity-symbol Recurringappointmentmaster pa-pd' },
    'socialactivity': { type: 'symbol', value: 'crmSymbolFont entity-symbol Socialactivity pa-pd' },
    'socialprofile': { type: 'symbol', value: 'crmSymbolFont entity-symbol Socialprofile pa-pd' },
    'kbarticle': { type: 'symbol', value: 'crmSymbolFont entity-symbol Kbarticle pa-pd' },
    'entitlement': { type: 'symbol', value: 'crmSymbolFont entity-symbol Entitlement pa-pd' },
    'bookableresource': { type: 'symbol', value: 'crmSymbolFont entity-symbol Bookableresource pa-pd' },
    'bookableresourcebooking': { type: 'symbol', value: 'crmSymbolFont entity-symbol Bookableresourcebooking pa-pd' },
    'msdyn_workorder': { type: 'symbol', value: 'crmSymbolFont entity-symbol msdyn_workorder pa-pd' },
    'msdyn_project': { type: 'symbol', value: 'crmSymbolFont entity-symbol msdyn_project pa-pd' }
};

/**
 * Helper function to get the icon definition for a table
 * @param tableName The logical name of the table
 * @param imageUrl Optional image URL from the icon cache
 * @returns The icon definition to use, or undefined if no icon available
 */
export function getTableIconDefinition(tableName: string, imageUrl?: string | null): ITableIconDefinition | undefined {
    const lowerTableName = tableName.toLowerCase();
    
    // First check if this is an OOB table with a symbol font icon
    if (OOB_TABLE_ICON_CLASSES[lowerTableName]) {
        return OOB_TABLE_ICON_CLASSES[lowerTableName];
    }
    
    // Otherwise, use the image URL if provided
    if (imageUrl) {
        return { type: 'image', value: imageUrl };
    }
    
    return undefined;
}
