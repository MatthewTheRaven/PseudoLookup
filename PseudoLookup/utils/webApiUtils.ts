/**
 * Web API utility functions for retrieving record data from Dataverse
 */

import { IInputs } from "../generated/ManifestTypes";
import { 
    ILookupRecord, 
    IParsedRecord, 
    IRecordRetrievalResult,
    ILookupFilter,
    ITableIconCache
} from "../types";
import { 
    getPrimaryNameField, 
    normalizeGuid 
} from "./dataUtils";

/**
 * Extended lookup options that includes additional properties
 * The standard PCF types may not include these but Dataverse supports them
 */
interface ExtendedLookupOptions extends ComponentFramework.UtilityApi.LookupOptions {
    filters?: ILookupFilter[];
    defaultEntityType?: string;
    disableMru?: boolean;
}

/**
 * Entity metadata response structure
 */
interface EntityMetadataResponse {
    IconVectorName?: string;
    IconSmallName?: string;
    PrimaryNameAttribute?: string;
    LogicalName?: string;
}

/**
 * Retrieves table icon URL from entity metadata
 * @param context The PCF context
 * @param tableName The logical name of the table
 * @returns Icon URL or null if not available
 */
export async function getTableIconUrl(
    context: ComponentFramework.Context<IInputs>,
    tableName: string
): Promise<string | null> {
    try {
        const utils = context.utils as {
            getEntityMetadata?: (entityName: string, attributes?: string[]) => Promise<EntityMetadataResponse>;
        };
        
        if (utils.getEntityMetadata) {
            const metadata = await utils.getEntityMetadata(tableName);
            
            if (metadata) {
                const iconName = metadata.IconVectorName ?? metadata.IconSmallName;
                
                if (iconName) {
                    const clientUrl = (context as unknown as { page?: { getClientUrl?: () => string } }).page?.getClientUrl?.() ?? '';
                    
                    if (clientUrl) {
                        return `${clientUrl}/WebResources/${iconName}`;
                    }
                }
            }
        }
        
        return null;
    } catch (error) {
        console.warn(`[PseudoLookup] Could not retrieve icon for table ${tableName}:`, error);
        return null;
    }
}

/**
 * Retrieves icon URLs for multiple tables
 * @param context The PCF context
 * @param tableNames Array of table logical names
 * @returns Cache object with table names as keys and icon URLs as values
 */
export async function getTableIconsForTables(
    context: ComponentFramework.Context<IInputs>,
    tableNames: string[]
): Promise<ITableIconCache> {
    const iconCache: ITableIconCache = {};
    
    if (!tableNames || tableNames.length === 0) {
        return iconCache;
    }
    
    // Fetch icons in parallel
    const iconPromises = tableNames.map(async (tableName) => {
        const iconUrl = await getTableIconUrl(context, tableName);
        return { tableName: tableName.toLowerCase(), iconUrl };
    });
    
    const results = await Promise.all(iconPromises);
    
    for (const { tableName, iconUrl } of results) {
        iconCache[tableName] = iconUrl;
    }
    
    return iconCache;
}

/**
 * Retrieves a single record from Dataverse using the Web API
 * @param context The PCF context with Web API access
 * @param tableName The logical name of the table
 * @param recordId The GUID of the record
 * @returns The record data with primary name, or null if not found/accessible
 */
export async function retrieveRecord(
    context: ComponentFramework.Context<IInputs>,
    tableName: string,
    recordId: string
): Promise<ILookupRecord | null> {
    try {
        const primaryNameField = getPrimaryNameField(tableName);
        const normalizedId = normalizeGuid(recordId);
        
        const result = await context.webAPI.retrieveRecord(
            tableName,
            normalizedId,
            `?$select=${primaryNameField}`
        );
        
        if (result) {
            const resultObj = result as Record<string, string | undefined>;
            const name = resultObj[primaryNameField] ?? 
                         resultObj.name ?? 
                         resultObj.fullname ?? 
                         resultObj.subject ?? 
                         resultObj.title ??
                         'Unnamed Record';
            
            return {
                id: normalizedId,
                table: tableName.toLowerCase(),
                name: name
            };
        }
        
        return null;
    } catch (error: unknown) {
        // Handle specific error codes silently (deleted/no permission)
        if (isWebApiError(error)) {
            const errorCode = error.code ?? error.errorCode;
            const errorMessage = error.message ?? '';
            
            if (errorCode === 404 || errorCode === '0x80040217' || errorMessage.includes('does not exist') ||
                errorCode === 403 || errorCode === '0x80040220') {
                return null;
            }
        }
        
        console.error(`[PseudoLookup] Error retrieving record ${recordId} from ${tableName}:`, error);
        return null;
    }
}

/**
 * Type guard to check if error is a Web API error with code
 */
function isWebApiError(error: unknown): error is { code?: number | string; errorCode?: number | string; message?: string } {
    return typeof error === 'object' && error !== null;
}

/**
 * Retrieves all records from the parsed bound value and returns valid records
 * Handles deleted records by removing them from the result
 * @param context The PCF context with Web API access
 * @param parsedRecords Array of parsed records (id and table only)
 * @returns Object containing valid records and information about deleted records
 */
export async function retrieveAllRecords(
    context: ComponentFramework.Context<IInputs>,
    parsedRecords: IParsedRecord[]
): Promise<IRecordRetrievalResult> {
    const validRecords: ILookupRecord[] = [];
    const deletedRecordIds: string[] = [];
    
    if (!parsedRecords || parsedRecords.length === 0) {
        return {
            validRecords,
            deletedRecordIds,
            hasDeletedRecords: false
        };
    }
    
    // Process all records in parallel
    const retrievalPromises = parsedRecords.map(async (parsed) => {
        const record = await retrieveRecord(context, parsed.table, parsed.id);
        return { parsed, record };
    });
    
    const results = await Promise.all(retrievalPromises);
    
    for (const { parsed, record } of results) {
        if (record) {
            validRecords.push(record);
        } else {
            deletedRecordIds.push(parsed.id);
        }
    }
    
    return {
        validRecords,
        deletedRecordIds,
        hasDeletedRecords: deletedRecordIds.length > 0
    };
}

/**
 * Retrieves table metadata to get the primary name attribute
 * This is a fallback when common mappings don't work
 * @param context The PCF context
 * @param tableName The logical name of the table
 * @returns The primary name attribute, or null if not retrievable
 */
export async function getTablePrimaryNameAttribute(
    context: ComponentFramework.Context<IInputs>,
    tableName: string
): Promise<string | null> {
    try {
        // Try to use entity metadata if available
        // Note: This may not be available in all PCF contexts
        const utils = context.utils as {
            getEntityMetadata?: (entityName: string, attributes?: string[]) => Promise<{
                PrimaryNameAttribute?: string;
            }>;
        };
        
        if (utils.getEntityMetadata) {
            const metadata = await utils.getEntityMetadata(tableName);
            if (metadata?.PrimaryNameAttribute) {
                return metadata.PrimaryNameAttribute;
            }
        }
    } catch (error) {
        console.warn(`Could not retrieve metadata for table ${tableName}:`, error);
    }
    
    return null;
}

/**
 * Opens the lookup dialog using Xrm.Utility.lookupObjects
 * @param context The PCF context
 * @param entityTypes Array of table logical names
 * @param allowMultiSelect Whether multiple selection is allowed
 * @param filters Array of FetchXML filters to apply
 * @param defaultEntityType The default table to show in the lookup dialog
 * @param disableMru Whether to disable Most Recently Used items
 * @returns Array of selected lookup results, or empty array if cancelled
 */
export async function openLookupDialog(
    context: ComponentFramework.Context<IInputs>,
    entityTypes: string[],
    allowMultiSelect: boolean,
    filters?: ILookupFilter[],
    defaultEntityType?: string,
    disableMru?: boolean
): Promise<ComponentFramework.LookupValue[]> {
    try {
        const lookupOptions: ExtendedLookupOptions = {
            entityTypes,
            allowMultiSelect
        };
        
        // Add filters if provided
        if (filters && filters.length > 0) {
            lookupOptions.filters = filters;
        }
        
        // Set defaultEntityType
        if (defaultEntityType) {
            lookupOptions.defaultEntityType = defaultEntityType;
        }
        
        // Set disableMru
        if (disableMru !== undefined) {
            lookupOptions.disableMru = disableMru;
        }
        
        const result = await context.utils.lookupObjects(lookupOptions);
        
        return result || [];
    } catch (error) {
        // User cancelled the dialog - this is not an error
        if (isUserCancelledError(error)) {
            return [];
        }
        
        console.error('Error opening lookup dialog:', error);
        throw error;
    }
}

/**
 * Checks if the error indicates the user cancelled the lookup dialog
 */
function isUserCancelledError(error: unknown): boolean {
    if (typeof error === 'object' && error !== null) {
        const err = error as { code?: number; message?: string };
        // Common cancellation error codes/messages
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
 * Opens a record form in a modal dialog or navigates to it
 * @param context The PCF context
 * @param entityName The logical name of the table
 * @param entityId The GUID of the record
 * @param openInModal Whether to open in a modal dialog (true) or navigate (false)
 * @returns Promise that resolves when the form is closed
 */
export async function openRecordForm(
    context: ComponentFramework.Context<IInputs>,
    entityName: string,
    entityId: string,
    openInModal = false
): Promise<void> {
    try {
        await context.navigation.openForm({
            entityName,
            entityId: normalizeGuid(entityId),
            openInNewWindow: false,
            windowPosition: openInModal ? 2 : 1 // 2 = Modal dialog (center), 1 = Inline (navigate)
        });
    } catch (error) {
        console.error(`Error opening form for ${entityName} record ${entityId}:`, error);
        throw error;
    }
}
