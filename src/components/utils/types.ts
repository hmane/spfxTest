// types.ts - Simplified and enhanced version

// Core enums with better organization
export type RenderMode = 'modal' | 'samepage' | 'newtab';
export type PickerMode = 'libraryFirst' | 'contentTypeFirst' | 'mixed';
export type OverwritePolicy = 'overwrite' | 'skip' | 'suffix';
export type UploadSelectionScope = 'single' | 'multiple';

// Simplified library configuration
export interface LibraryOption {
	/** Server-relative URL to the library root */
	serverRelativeUrl: string;
	/** Friendly display name */
	label?: string;
	/** Default subfolder for uploads */
	defaultFolder?: string;
	/** Minimal view ID for bulk editing (optional) */
	minimalViewId?: string;
	/** Allowed content types - 'all' or array of IDs */
	allowedContentTypeIds?: string[] | 'all';
}

// Content type information
export interface ContentTypeInfo {
	id: string;
	name: string;
	description?: string;
	hidden?: boolean;
}

// User selection from pickers
export interface DestinationChoice {
	libraryUrl: string;
	contentTypeId?: string;
	folderPath?: string;
	// Friendly names for display
	libraryTitle?: string;
	contentTypeName?: string;
}

// Upload results with better error handling
export interface UploadBatchResult {
	/** Successfully uploaded item IDs */
	itemIds: number[];
	/** Files that failed with error details */
	failed: Array<{ name: string; message: string; error?: any }>;
	/** Files that were skipped (already existed) */
	skipped?: string[];
}

// Simplified file progress tracking
export interface FileProgress {
	fileName: string;
	percent: number;
	status: 'queued' | 'uploading' | 'done' | 'error' | 'skipped';
	errorMessage?: string;
	itemId?: number;
}

// Extended file state for internal component use
export interface FileUploadState extends FileProgress {
	file: File;
	attempts: number;
	targetFileName?: string;
}

// Web part configuration (streamlined)
export interface UploadAndEditWebPartProps {
	// Core behavior
	pickerMode: PickerMode;
	renderMode: RenderMode;
	selectionScope: UploadSelectionScope;
	showContentTypePicker: boolean;
	overwritePolicy: OverwritePolicy;

	// Library configuration
	libraries: LibraryOption[];
	globalAllowedContentTypeIds?: string[] | 'all';

	// Editor behavior
	enableBulkAutoRefresh: boolean;
	bulkWatchAllItems: boolean;

	// UI customization
	buttonLabel?: string;
	dropzoneHint?: string;
	successToast?: string;

	// Advanced options
	disableDomNudges: boolean;
	sandboxExtra?: string;
}

// Service interface (simplified)
export interface SharePointService {
	getLibraryTitle(libraryUrl: string): Promise<string>;
	getLibraryContentTypes(libraryUrl: string): Promise<ContentTypeInfo[]>;

	fileExists(
		libraryUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean>;

	uploadFileWithProgress(
		libraryUrl: string,
		folderPath: string | undefined,
		file: File,
		onProgress: (percent: number) => void,
		overwritePolicy: OverwritePolicy,
		chunkSizeBytes?: number,
		confirmOverwrite?: (fileName: string) => Promise<boolean>
	): Promise<{ itemId: number; serverRelativeUrl: string; uniqueId: string }>;

	setItemContentType(libraryUrl: string, itemId: number, contentTypeId: string): Promise<void>;
}

// Editor launcher events
export interface LauncherDeterminedInfo {
	mode: 'single' | 'bulk';
	url: string;
}

export interface LauncherOpenInfo {
	mode: 'single' | 'bulk';
	url: string;
}

// Error handling
export interface SPFormError {
	message: string;
	code?: string;
	context?: any;
	originalError?: any;
}

// Component state management
export type ComponentStage = 'idle' | 'destination' | 'upload' | 'editing';

export interface ComponentState {
	stage: ComponentStage;
	dialogOpen: boolean;
	choice?: DestinationChoice;
	pendingFiles: File[];
	uploadedItemIds: number[];
	errorMsg: string | null;
	isLoading: boolean;
	loadingMessage?: string;
}

// Utility types for better type safety
export type NonEmptyArray<T> = [T, ...T[]];

export interface UploadOptions {
	allowMultiple: boolean;
	acceptedFileTypes?: string[];
	maxFileSize?: number;
	maxFiles?: number;
}

// Event callbacks
export interface ComponentCallbacks {
	onFilesPicked?: (files: File[]) => void;
	onUploadStart?: (files: File[]) => void;
	onUploadProgress?: (progress: FileProgress[]) => void;
	onUploadComplete?: (result: UploadBatchResult) => void;
	onEditingStart?: (itemIds: number[]) => void;
	onEditingComplete?: () => void;
	onError?: (error: SPFormError) => void;
}

// Hook return types for better reusability
export interface UseUploadState {
	state: ComponentState;
	actions: {
		setStage: (stage: ComponentStage) => void;
		setDialogOpen: (open: boolean) => void;
		setChoice: (choice: DestinationChoice | undefined) => void;
		setPendingFiles: (files: File[]) => void;
		setUploadedItemIds: (ids: number[]) => void;
		setError: (error: string | null) => void;
		setLoading: (loading: boolean, message?: string) => void;
		reset: () => void;
	};
}

// Configuration validation
export interface ConfigValidationResult {
	isValid: boolean;
	errors: string[];
	warnings: string[];
}

// Toast notification types
export type ToastType = 'success' | 'error' | 'warning' | 'info';

export interface ToastMessage {
	id?: string;
	type: ToastType;
	title?: string;
	message: string;
	timeout?: number;
	persistent?: boolean;
}
