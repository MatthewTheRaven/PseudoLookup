/**
 * LookupBadge component - Displays a single selected record as a badge
 * Supports optional navigation on click and always allows deletion
 */

import * as React from 'react';
import { ILookupRecord, ITableIconDefinition, getTableIconDefinition } from '../types';

export interface ILookupBadgeProps {
    /** The record data to display */
    record: ILookupRecord;
    /** Icon URL for the table from cache (optional) */
    iconUrl?: string;
    /** Whether clicking the badge navigates to the record */
    enableLinking: boolean;
    /** Whether the control is disabled/read-only */
    isDisabled: boolean;
    /** Whether this specific record is still loading */
    isLoading?: boolean;
    /** Callback when delete button is clicked */
    onRemove: (recordId: string) => void;
    /** Callback when badge text is clicked (for navigation) */
    onNavigate: (record: ILookupRecord) => void;
}

/**
 * LookupBadge displays a selected record with optional navigation and delete functionality
 */
export const LookupBadge: React.FC<ILookupBadgeProps> = ({
    record,
    iconUrl,
    enableLinking,
    isDisabled,
    isLoading,
    onRemove,
    onNavigate
}) => {
    // State for icon loading error (only applicable for image icons)
    const [iconError, setIconError] = React.useState(false);
    
    // Get the icon definition for this record's table
    const iconDefinition: ITableIconDefinition | undefined = React.useMemo(() => {
        if (isLoading) return undefined;
        return getTableIconDefinition(record.table, iconUrl);
    }, [record.table, iconUrl, isLoading]);
    
    /**
     * Handles icon load error - hides the icon
     */
    const handleIconError = React.useCallback(() => {
        setIconError(true);
    }, []);
    
    /**
     * Handles click on the badge text area
     * Triggers navigation if linking is enabled (even when read-only/disabled)
     */
    const handleTextClick = React.useCallback((event: React.MouseEvent<HTMLSpanElement>) => {
        event.preventDefault();
        event.stopPropagation();
        
        if (enableLinking) {
            onNavigate(record);
        }
    }, [enableLinking, onNavigate, record]);

    /**
     * Handles keyboard interaction on the badge text
     * Supports Enter and Space for accessibility
     */
    const handleTextKeyDown = React.useCallback((event: React.KeyboardEvent<HTMLSpanElement>) => {
        if (enableLinking && (event.key === 'Enter' || event.key === ' ')) {
            event.preventDefault();
            event.stopPropagation();
            onNavigate(record);
        }
    }, [enableLinking, onNavigate, record]);

    /**
     * Handles click on the delete button
     * Stops propagation to prevent triggering badge click
     */
    const handleDeleteClick = React.useCallback((event: React.MouseEvent<HTMLButtonElement>) => {
        event.preventDefault();
        event.stopPropagation();
        
        if (!isDisabled) {
            onRemove(record.id);
        }
    }, [isDisabled, onRemove, record.id]);

    /**
     * Handles keyboard interaction on delete button
     */
    const handleDeleteKeyDown = React.useCallback((event: React.KeyboardEvent<HTMLButtonElement>) => {
        if (!isDisabled && (event.key === 'Enter' || event.key === ' ')) {
            event.preventDefault();
            event.stopPropagation();
            onRemove(record.id);
        }
    }, [isDisabled, onRemove, record.id]);

    const badgeClasses = ['lookup-badge'];
    if (isDisabled) {
        badgeClasses.push('disabled');
    }
    if (isLoading) {
        badgeClasses.push('loading');
    }

    const textClasses = ['lookup-badge-text'];
    // Enable linking even when disabled (read-only), just not when loading
    if (enableLinking && !isLoading) {
        textClasses.push('clickable');
    }
    
    // Determine what icon to show (if any)
    const showSymbolIcon = !isLoading && iconDefinition?.type === 'symbol';
    const showImageIcon = !isLoading && iconDefinition?.type === 'image' && !iconError;

    return (
        <div 
            className={badgeClasses.join(' ')}
            role="listitem"
            aria-label={isLoading ? `Loading record from ${record.table}` : `Selected record: ${record.name} from ${record.table}`}
        >
            {/* Loading spinner - shown when record is loading */}
            {isLoading && (
                <div 
                    className="lookup-badge-spinner" 
                    role="progressbar" 
                    aria-label="Loading record"
                />
            )}
            
            {/* Symbol font icon - for OOB tables */}
            {showSymbolIcon && (
                <div className="lookup-badge-icon-wrapper" aria-hidden="true">
                    <span className={iconDefinition.value} />
                </div>
            )}
            
            {/* Image icon - for custom tables with image URLs */}
            {showImageIcon && (
                <img
                    className="lookup-badge-icon"
                    src={iconDefinition.value}
                    alt=""
                    aria-hidden="true"
                    onError={handleIconError}
                />
            )}
            
            {/* Badge text - clickable when linking is enabled (even in read-only mode) */}
            <span
                className={textClasses.join(' ')}
                onClick={isLoading ? undefined : handleTextClick}
                onKeyDown={isLoading ? undefined : handleTextKeyDown}
                role={enableLinking && !isLoading ? 'link' : undefined}
                tabIndex={enableLinking && !isLoading ? 0 : -1}
                aria-label={isLoading ? 'Loading...' : (enableLinking ? `Open ${record.name}` : record.name)}
                title={isLoading ? 'Loading...' : record.name}
            >
                {isLoading ? 'Loading...' : record.name}
            </span>
            
            {/* Delete button - hidden when disabled (read-only) or loading */}
            {!isDisabled && !isLoading && (
                <button
                    className="lookup-badge-delete"
                    onClick={handleDeleteClick}
                    onKeyDown={handleDeleteKeyDown}
                    type="button"
                    aria-label={`Remove ${record.name}`}
                    title={`Remove ${record.name}`}
                >
                    <span 
                        className="symbolFont Cancel-symbol pa-pj pa-pk"
                        aria-hidden="true"
                    />
                </button>
            )}
        </div>
    );
};

export default LookupBadge;
