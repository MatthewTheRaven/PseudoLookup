/**
 * Utility exports for the Multi-Table Lookup control
 */

export {
    isValidGuid,
    normalizeGuid,
    parseBoundValue,
    serializeToBoundValue,
    parseLookupTables,
    getPrimaryKeyField,
    getPrimaryNameField,
    getPluralTableName,
    generateExclusionFilter,
    generateAllExclusionFilters,
    isRecordSelected,
    removeRecordById,
    escapeXml
} from './dataUtils';

export {
    retrieveRecord,
    retrieveAllRecords,
    getTablePrimaryNameAttribute,
    openLookupDialog,
    openRecordForm,
    getTableIconUrl,
    getTableIconsForTables
} from './webApiUtils';
