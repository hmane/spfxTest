import { SPFxContext } from './pnpConfig';

/**
 * SharePoint field types enum
 */
export enum SPFieldType {
	Text = 'Text',
	Note = 'Note',
	Number = 'Number',
	Currency = 'Currency',
	DateTime = 'DateTime',
	Boolean = 'Boolean',
	Choice = 'Choice',
	MultiChoice = 'MultiChoice',
	Lookup = 'Lookup',
	User = 'User',
	URL = 'URL',
	Calculated = 'Calculated',
	Attachments = 'Attachments',
	Guid = 'Guid',
	Integer = 'Integer',
	Counter = 'Counter',
	TaxonomyFieldType = 'TaxonomyFieldType',
	TaxonomyFieldTypeMulti = 'TaxonomyFieldTypeMulti',
}

/**
 * Common SharePoint REST API query parameters
 */
export interface IQueryParams {
	select?: string[];
	expand?: string[];
	filter?: string;
	orderBy?: string[];
	top?: number;
	skip?: number;
}

/**
 * Build OData query string from parameters
 */
export const buildODataQuery = (params: IQueryParams): string => {
	const queryParts: string[] = [];

	if (params.select && params.select.length > 0) {
		queryParts.push(`$select=${params.select.join(',')}`);
	}

	if (params.expand && params.expand.length > 0) {
		queryParts.push(`$expand=${params.expand.join(',')}`);
	}

	if (params.filter) {
		queryParts.push(`$filter=${encodeURIComponent(params.filter)}`);
	}

	if (params.orderBy && params.orderBy.length > 0) {
		queryParts.push(`$orderby=${params.orderBy.join(',')}`);
	}

	if (params.top !== undefined) {
		queryParts.push(`$top=${params.top}`);
	}

	if (params.skip !== undefined) {
		queryParts.push(`$skip=${params.skip}`);
	}

	return queryParts.length > 0 ? `?${queryParts.join('&')}` : '';
};

/**
 * Format SharePoint date for REST API
 */
export const formatDateForSP = (date: Date): string => {
	return date.toISOString();
};

/**
 * Parse SharePoint date string to Date object
 */
export const parseSPDate = (dateString: string): Date => {
	return new Date(dateString);
};

/**
 * Format file size in human readable format
 */
export const formatFileSize = (bytes: number): string => {
	if (bytes === 0) return '0 Bytes';

	const k = 1024;
	const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
	const i = Math.floor(Math.log(bytes) / Math.log(k));

	return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

/**
 * Extract file extension from filename
 */
export const getFileExtension = (filename: string): string => {
	const lastDot = filename.lastIndexOf('.');
	return lastDot > 0 ? filename.substring(lastDot + 1).toLowerCase() : '';
};

/**
 * Check if file is an image based on extension
 */
export const isImageFile = (filename: string): boolean => {
	const imageExtensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'];
	const extension = getFileExtension(filename);
	return imageExtensions.includes(extension);
};

/**
 * Check if file is a document based on extension
 */
export const isDocumentFile = (filename: string): boolean => {
	const docExtensions = ['doc', 'docx', 'pdf', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'rtf'];
	const extension = getFileExtension(filename);
	return docExtensions.includes(extension);
};

/**
 * Generate GUID (similar to SharePoint GUID format)
 */
export const generateGuid = (): string => {
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
		const r = (Math.random() * 16) | 0;
		const v = c === 'x' ? r : (r & 0x3) | 0x8;
		return v.toString(16);
	});
};

/**
 * Sanitize filename for SharePoint (remove invalid characters)
 */
export const sanitizeFileName = (filename: string): string => {
	// SharePoint invalid characters: \ / : * ? " < > | # { } % ~ &
	return filename.replace(/[\\/:*?"<>|#{}\%~&]/g, '_');
};

/**
 * Create SharePoint list item entity type name
 */
export const getListItemEntityTypeName = (listTitle: string): string => {
	return `SP.Data.${listTitle.replace(/\s/g, '_x0020_')}ListItem`;
};

/**
 * Build SharePoint REST API URL
 */
export const buildSPRestUrl = (
	context: SPFxContext,
	endpoint: string,
	params?: IQueryParams
): string => {
	const baseUrl = `${context.pageContext.web.absoluteUrl}/_api/${endpoint}`;
	const queryString = params ? buildODataQuery(params) : '';
	return `${baseUrl}${queryString}`;
};

/**
 * Handle SharePoint REST API errors
 */
export const handleSPError = (error: any): Error => {
	let message = 'An unknown SharePoint error occurred';

	if (error?.response?.data?.error?.message?.value) {
		message = error.response.data.error.message.value;
	} else if (error?.message) {
		message = error.message;
	} else if (typeof error === 'string') {
		message = error;
	}

	return new Error(`SharePoint Error: ${message}`);
};

/**
 * Delay function for throttling requests
 */
export const delay = (ms: number): Promise<void> => {
	return new Promise((resolve) => setTimeout(resolve, ms));
};

/**
 * Retry function with exponential backoff
 */
export const retryWithBackoff = async <T>(
	operation: () => Promise<T>,
	maxRetries: number = 3,
	baseDelay: number = 1000
): Promise<T> => {
	let lastError: Error;

	for (let attempt = 0; attempt <= maxRetries; attempt++) {
		try {
			return await operation();
		} catch (error) {
			lastError = error instanceof Error ? error : new Error(String(error));

			if (attempt === maxRetries) {
				throw lastError;
			}

			const delayMs = baseDelay * Math.pow(2, attempt);
			console.warn(`Attempt ${attempt + 1} failed, retrying in ${delayMs}ms...`, error);
			await delay(delayMs);
		}
	}

	throw lastError!;
};
