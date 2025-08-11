/**
 * Base error class for the component library
 */
export class SPFxComponentLibraryError extends Error {
	public readonly code: string;
	public readonly context?: any;

	constructor(message: string, code: string = 'UNKNOWN_ERROR', context?: any) {
		super(message);
		this.name = 'SPFxComponentLibraryError';
		this.code = code;
		this.context = context;
	}
}

/**
 * SharePoint specific error
 */
export class SharePointError extends SPFxComponentLibraryError {
	constructor(message: string, code: string = 'SHAREPOINT_ERROR', context?: any) {
		super(message, code, context);
		this.name = 'SharePointError';
	}
}

/**
 * Microsoft Graph specific error
 */
export class GraphError extends SPFxComponentLibraryError {
	constructor(message: string, code: string = 'GRAPH_ERROR', context?: any) {
		super(message, code, context);
		this.name = 'GraphError';
	}
}

/**
 * Error codes enum
 */
export enum ErrorCodes {
	// General errors
	UNKNOWN_ERROR = 'UNKNOWN_ERROR',
	VALIDATION_ERROR = 'VALIDATION_ERROR',
	NETWORK_ERROR = 'NETWORK_ERROR',

	// SharePoint errors
	SHAREPOINT_LIST_NOT_FOUND = 'SHAREPOINT_LIST_NOT_FOUND',
	SHAREPOINT_ITEM_NOT_FOUND = 'SHAREPOINT_ITEM_NOT_FOUND',
	SHAREPOINT_ACCESS_DENIED = 'SHAREPOINT_ACCESS_DENIED',
	SHAREPOINT_FIELD_NOT_FOUND = 'SHAREPOINT_FIELD_NOT_FOUND',

	// Graph errors
	GRAPH_USER_NOT_FOUND = 'GRAPH_USER_NOT_FOUND',
	GRAPH_ACCESS_DENIED = 'GRAPH_ACCESS_DENIED',
	GRAPH_THROTTLED = 'GRAPH_THROTTLED',

	// Component errors
	COMPONENT_NOT_INITIALIZED = 'COMPONENT_NOT_INITIALIZED',
	INVALID_CONTEXT = 'INVALID_CONTEXT',
	INVALID_PROPS = 'INVALID_PROPS',
}

/**
 * Parse error and return standardized error object
 */
export const parseError = (error: any): SPFxComponentLibraryError => {
	if (error instanceof SPFxComponentLibraryError) {
		return error;
	}

	let message = 'An unknown error occurred';
	let code = ErrorCodes.UNKNOWN_ERROR;
	let context: any = null;

	// Handle different error types
	if (error?.response) {
		// HTTP error response
		const response = error.response;
		context = { status: response.status, statusText: response.statusText };

		if (response.status === 404) {
			code = ErrorCodes.SHAREPOINT_ITEM_NOT_FOUND;
			message = 'The requested item was not found';
		} else if (response.status === 403) {
			code = ErrorCodes.SHAREPOINT_ACCESS_DENIED;
			message = 'Access denied';
		} else if (response.status === 429) {
			code = ErrorCodes.GRAPH_THROTTLED;
			message = 'Request was throttled';
		}

		// Try to extract detailed error message
		if (response.data?.error?.message) {
			message = response.data.error.message;
		} else if (response.data?.message) {
			message = response.data.message;
		}
	} else if (error?.message) {
		message = error.message;
	} else if (typeof error === 'string') {
		message = error;
	}

	return new SPFxComponentLibraryError(message, code, context);
};

/**
 * Log error with context
 */
export const logError = (error: any, context?: string): void => {
	const parsedError = parseError(error);

	console.error(
		`[SPFx Component Library${context ? ` - ${context}` : ''}] ${parsedError.code}: ${
			parsedError.message
		}`,
		{
			error: parsedError,
			context: parsedError.context,
			stack: parsedError.stack,
		}
	);
};

/**
 * Handle async operation with error handling
 */
export const handleAsyncOperation = async <T>(
	operation: () => Promise<T>,
	context?: string
): Promise<T | null> => {
	try {
		return await operation();
	} catch (error) {
		logError(error, context);
		return null;
	}
};

/**
 * Validate required parameters
 */
export const validateRequired = (params: Record<string, any>, requiredFields: string[]): void => {
	const missingFields = requiredFields.filter(
		(field) => params[field] === null || params[field] === undefined || params[field] === ''
	);

	if (missingFields.length > 0) {
		throw new SPFxComponentLibraryError(
			`Missing required parameters: ${missingFields.join(', ')}`,
			ErrorCodes.VALIDATION_ERROR,
			{ missingFields }
		);
	}
};
