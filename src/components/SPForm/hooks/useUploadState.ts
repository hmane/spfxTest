// hooks/useUploadState.ts - Centralized state management
import { useState, useCallback, useReducer, useMemo } from 'react';
import {
	DestinationChoice,
	UseUploadState,
	SPFormError,
	ComponentCallbacks,
	ConfigValidationResult,
    ComponentStage,
    ComponentState,
} from '../../utils/types';


// Action types for reducer
type StateAction =
	| { type: 'SET_STAGE'; payload: ComponentStage }
	| { type: 'SET_DIALOG_OPEN'; payload: boolean }
	| { type: 'SET_CHOICE'; payload: DestinationChoice | undefined }
	| { type: 'SET_PENDING_FILES'; payload: File[] }
	| { type: 'SET_UPLOADED_ITEM_IDS'; payload: number[] }
	| { type: 'SET_ERROR'; payload: string | null }
	| { type: 'SET_LOADING'; payload: { loading: boolean; message?: string } }
	| { type: 'RESET' };

// Initial state
const initialState: ComponentState = {
	stage: 'idle',
	dialogOpen: false,
	choice: undefined,
	pendingFiles: [],
	uploadedItemIds: [],
	errorMsg: null,
	isLoading: false,
	loadingMessage: undefined,
};

// State reducer
function stateReducer(state: ComponentState, action: StateAction): ComponentState {
	switch (action.type) {
		case 'SET_STAGE':
			return { ...state, stage: action.payload };

		case 'SET_DIALOG_OPEN':
			return { ...state, dialogOpen: action.payload };

		case 'SET_CHOICE':
			return { ...state, choice: action.payload };

		case 'SET_PENDING_FILES':
			return { ...state, pendingFiles: action.payload };

		case 'SET_UPLOADED_ITEM_IDS':
			return { ...state, uploadedItemIds: action.payload };

		case 'SET_ERROR':
			return { ...state, errorMsg: action.payload, isLoading: false };

		case 'SET_LOADING':
			return {
				...state,
				isLoading: action.payload.loading,
				loadingMessage: action.payload.message,
			};

		case 'RESET':
			return initialState;

		default:
			return state;
	}
}

// Custom hook for upload state management
export function useUploadState(callbacks?: ComponentCallbacks): UseUploadState {
	const [state, dispatch] = useReducer(stateReducer, initialState);

	// Action creators with callback integration
	const actions = useMemo(
		() => ({
			setStage: (stage: ComponentStage) => {
				dispatch({ type: 'SET_STAGE', payload: stage });
			},

			setDialogOpen: (open: boolean) => {
				dispatch({ type: 'SET_DIALOG_OPEN', payload: open });
			},

			setChoice: (choice: DestinationChoice | undefined) => {
				dispatch({ type: 'SET_CHOICE', payload: choice });
			},

			setPendingFiles: (files: File[]) => {
				dispatch({ type: 'SET_PENDING_FILES', payload: files });
				callbacks?.onFilesPicked?.(files);
			},

			setUploadedItemIds: (ids: number[]) => {
				dispatch({ type: 'SET_UPLOADED_ITEM_IDS', payload: ids });
			},

			setError: (error: string | null) => {
				dispatch({ type: 'SET_ERROR', payload: error });
				if (error && callbacks?.onError) {
					callbacks.onError({
						message: error,
						code: 'COMPONENT_ERROR',
					});
				}
			},

			setLoading: (loading: boolean, message?: string) => {
				dispatch({ type: 'SET_LOADING', payload: { loading, message } });
			},

			reset: () => {
				dispatch({ type: 'RESET' });
			},
		}),
		[callbacks]
	);

	return { state, actions };
}

// Custom hook for file validation
export function useFileValidation(options?: {
	maxFileSize?: number;
	allowedExtensions?: string[];
	maxFiles?: number;
}) {
	const validateFiles = useCallback(
		(files: File[]): { valid: File[]; invalid: Array<{ file: File; reason: string }> } => {
			const valid: File[] = [];
			const invalid: Array<{ file: File; reason: string }> = [];

			files.forEach((file) => {
				// Check file size
				if (options?.maxFileSize && file.size > options.maxFileSize) {
					invalid.push({
						file,
						reason: `File too large (${formatFileSize(file.size)} > ${formatFileSize(
							options.maxFileSize
						)})`,
					});
					return;
				}

				// Check file extension
				if (options?.allowedExtensions?.length) {
					const extension = getFileExtension(file.name);
					if (!options.allowedExtensions.includes(extension)) {
						invalid.push({
							file,
							reason: `File type not allowed (${extension})`,
						});
						return;
					}
				}

				// Check for empty files
				if (file.size === 0) {
					invalid.push({
						file,
						reason: 'Empty file',
					});
					return;
				}

				valid.push(file);
			});

			// Check total file count
			if (options?.maxFiles && valid.length > options.maxFiles) {
				const excess = valid.splice(options.maxFiles);
				excess.forEach((file) => {
					invalid.push({
						file,
						reason: `Too many files (limit: ${options.maxFiles})`,
					});
				});
			}

			return { valid, invalid };
		},
		[options]
	);

	return { validateFiles };
}

// Utility functions
function formatFileSize(bytes: number): string {
	if (bytes === 0) return '0 Bytes';
	const k = 1024;
	const sizes = ['Bytes', 'KB', 'MB', 'GB'];
	const i = Math.floor(Math.log(bytes) / Math.log(k));
	return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function getFileExtension(filename: string): string {
	const lastDot = filename.lastIndexOf('.');
	return lastDot > 0 ? filename.substring(lastDot + 1).toLowerCase() : '';
}

// Hook for error handling with retry logic
export function useErrorHandler() {
	const [retryCount, setRetryCount] = useState(0);
	const maxRetries = 3;

	const handleError = useCallback((error: any, context?: string): SPFormError => {
		const formattedError: SPFormError = {
			message: error?.message || 'An unknown error occurred',
			code: error?.code || 'UNKNOWN_ERROR',
			context: context || 'Unknown context',
			originalError: error,
		};

		console.error(`[SPForm${context ? ` - ${context}` : ''}]:`, formattedError);
		return formattedError;
	}, []);

	const canRetry = useCallback(() => {
		return retryCount < maxRetries;
	}, [retryCount]);

	const retry = useCallback(() => {
		if (canRetry()) {
			setRetryCount((prev) => prev + 1);
			return true;
		}
		return false;
	}, [canRetry]);

	const resetRetry = useCallback(() => {
		setRetryCount(0);
	}, []);

	return {
		handleError,
		canRetry,
		retry,
		resetRetry,
		retryCount,
		maxRetries,
	};
}

// Hook for configuration validation
export function useConfigValidation(config: any): ConfigValidationResult {
	const validation = useMemo(() => {
		const errors: string[] = [];
		const warnings: string[] = [];

		// Validate required properties
		if (!config.libraries || !Array.isArray(config.libraries) || config.libraries.length === 0) {
			errors.push('At least one library must be configured');
		}

		// Validate library configurations
		config.libraries?.forEach((lib: any, index: number) => {
			if (!lib.serverRelativeUrl) {
				errors.push(`Library ${index + 1}: serverRelativeUrl is required`);
			}
			if (lib.serverRelativeUrl && !lib.serverRelativeUrl.startsWith('/')) {
				warnings.push(`Library ${index + 1}: serverRelativeUrl should start with '/'`);
			}
		});

		// Validate picker mode
		const validPickerModes = ['libraryFirst', 'contentTypeFirst', 'mixed'];
		if (config.pickerMode && !validPickerModes.includes(config.pickerMode)) {
			errors.push(`Invalid picker mode: ${config.pickerMode}`);
		}

		// Validate render mode
		const validRenderModes = ['modal', 'samepage', 'newtab'];
		if (config.renderMode && !validRenderModes.includes(config.renderMode)) {
			errors.push(`Invalid render mode: ${config.renderMode}`);
		}

		return {
			isValid: errors.length === 0,
			errors,
			warnings,
		};
	}, [config]);

	return validation;
}
