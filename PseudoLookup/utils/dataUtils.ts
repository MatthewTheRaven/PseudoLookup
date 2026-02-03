/**
 * Utility functions for data parsing, serialization, and FetchXML generation
 */

import { 
    ILookupRecord, 
    IParsedRecord, 
    ILookupFilter, 
    COMMON_PRIMARY_NAME_FIELDS 
} from "../types";

/**
 * Validates if a string is a valid GUID format
 * @param value The string to validate
 * @returns True if valid GUID
 */
export function isValidGuid(value: string): boolean {
    if (!value) return false;
    // Support both with and without braces
    const cleanGuid = value.replace(/[{}]/g, '');
    const guidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
    return guidRegex.test(cleanGuid);
}

/**
 * Normalizes a GUID by removing braces and converting to lowercase
 * @param guid The GUID to normalize
 * @returns Normalized GUID string
 */
export function normalizeGuid(guid: string): string {
    if (!guid) return '';
    return guid.replace(/[{}]/g, '').toLowerCase();
}

/**
 * Parses the bound field value into an array of partial lookup records (without names)
 * @param boundValue The semicolon-delimited string value
 * @returns Array of objects with id and table properties
 */
export function parseBoundValue(boundValue: string | null | undefined): IParsedRecord[] {
    if (!boundValue || boundValue.trim() === '') {
        return [];
    }

    const records: IParsedRecord[] = [];
    const entries = boundValue.split(';').filter(entry => entry.trim() !== '');

    for (const entry of entries) {
        const colonIndex = entry.indexOf(':');
        if (colonIndex === -1) continue;

        const id = entry.substring(0, colonIndex).trim();
        const table = entry.substring(colonIndex + 1).trim();
        
        // Validate GUID format and table name
        if (isValidGuid(id) && table.length > 0) {
            records.push({ 
                id: normalizeGuid(id), 
                table: table.toLowerCase() 
            });
        }
    }

    return records;
}

/**
 * Converts an array of lookup records to the bound field format
 * @param records Array of lookup records
 * @returns Semicolon-delimited string
 */
export function serializeToBoundValue(records: ILookupRecord[]): string {
    if (!records || records.length === 0) {
        return '';
    }

    return records
        .map(record => `${normalizeGuid(record.id)}:${record.table.toLowerCase()}`)
        .join(';');
}

/**
 * Parses comma-separated lookup tables string into array
 * @param tablesString Comma-separated table names
 * @returns Array of table logical names
 */
export function parseLookupTables(tablesString: string | null | undefined): string[] {
    if (!tablesString || tablesString.trim() === '') {
        return [];
    }

    return tablesString
        .split(',')
        .map(table => table.trim().toLowerCase())
        .filter(table => table.length > 0);
}

/**
 * Gets the primary key field name for a table
 * Convention: tablename + "id"
 * @param tableName The logical name of the table
 * @returns The primary key field name
 */
export function getPrimaryKeyField(tableName: string): string {
    return `${tableName.toLowerCase()}id`;
}

/**
 * Gets the likely primary name field for a table
 * Uses common patterns and falls back to 'name'
 * @param tableName The logical name of the table
 * @returns The primary name field name
 */
export function getPrimaryNameField(tableName: string): string {
    const lowerTable = tableName.toLowerCase();
    
    // Check common mappings first
    if (COMMON_PRIMARY_NAME_FIELDS[lowerTable]) {
        return COMMON_PRIMARY_NAME_FIELDS[lowerTable];
    }
    
    // Default fallback to 'name'
    return 'name';
}

/**
 * Gets the plural form of a table name for Web API requests
 * Handles common pluralization patterns
 * @param tableName The logical name of the table
 * @returns The pluralized table name for Web API
 */
export function getPluralTableName(tableName: string): string {
    const lowerTable = tableName.toLowerCase();
    
    // Handle common irregular plurals
    const irregularPlurals: Record<string, string> = {
        'opportunity': 'opportunities',
        'activity': 'activities',
        'entity': 'entities',
        'category': 'categories',
        'currency': 'currencies',
        'territory': 'territories',
        'query': 'queries',
        'policy': 'policies'
    };
    
    if (irregularPlurals[lowerTable]) {
        return irregularPlurals[lowerTable];
    }
    
    // Handle words ending in 's', 'x', 'z', 'ch', 'sh'
    if (/[sxz]$/.test(lowerTable) || /[cs]h$/.test(lowerTable)) {
        return `${lowerTable}es`;
    }
    
    // Handle words ending in 'y' (preceded by consonant)
    if (/[^aeiou]y$/.test(lowerTable)) {
        return `${lowerTable.slice(0, -1)}ies`;
    }
    
    // Default: add 's'
    return `${lowerTable}s`;
}

/**
 * Generates FetchXML filter to exclude already-selected records for a specific table
 * @param tableName The logical name of the table
 * @param selectedRecords Currently selected records
 * @returns Filter object or null if no filter needed
 */
export function generateExclusionFilter(
    tableName: string,
    selectedRecords: ILookupRecord[]
): ILookupFilter | null {
    // Get selected record IDs for this specific table
    const selectedIdsForTable = selectedRecords
        .filter(record => record.table.toLowerCase() === tableName.toLowerCase())
        .map(record => normalizeGuid(record.id));

    // No filter needed if no records selected for this table
    if (selectedIdsForTable.length === 0) {
        return null;
    }

    // Get the primary key field name
    const primaryKeyField = getPrimaryKeyField(tableName);

    // Build the FetchXML filter with not-in condition
    const valueElements = selectedIdsForTable
        .map(id => `<value>{${id}}</value>`)
        .join('');

    const filterXml = `<filter type="and"><condition attribute="${primaryKeyField}" operator="not-in">${valueElements}</condition></filter>`;

    return {
        filterXml,
        entityLogicalName: tableName.toLowerCase()
    };
}

/**
 * Generates all exclusion filters for the lookup dialog
 * @param lookupTables Array of table logical names
 * @param selectedRecords Currently selected records
 * @returns Array of filter objects
 */
export function generateAllExclusionFilters(
    lookupTables: string[],
    selectedRecords: ILookupRecord[]
): ILookupFilter[] {
    const filters: ILookupFilter[] = [];

    for (const tableName of lookupTables) {
        const filter = generateExclusionFilter(tableName, selectedRecords);
        if (filter) {
            filters.push(filter);
        }
    }

    return filters;
}

/**
 * Checks if a record already exists in the selected records array
 * @param recordId The record ID to check
 * @param selectedRecords The current selection
 * @returns True if record is already selected
 */
export function isRecordSelected(recordId: string, selectedRecords: ILookupRecord[]): boolean {
    const normalizedId = normalizeGuid(recordId);
    return selectedRecords.some(record => normalizeGuid(record.id) === normalizedId);
}

/**
 * Removes a record from the selected records array by ID
 * @param recordId The record ID to remove
 * @param selectedRecords The current selection
 * @returns New array without the specified record
 */
export function removeRecordById(recordId: string, selectedRecords: ILookupRecord[]): ILookupRecord[] {
    const normalizedId = normalizeGuid(recordId);
    return selectedRecords.filter(record => normalizeGuid(record.id) !== normalizedId);
}

/**
 * Escapes special XML characters in a string
 * @param str The string to escape
 * @returns XML-safe string
 */
export function escapeXml(str: string): string {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}
