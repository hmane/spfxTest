// src/components/SPForm/types.ts - Updated with missing types

// ---------- Shared enums / literals ----------
export type RenderMode = 'modal' | 'samepage' | 'newtab';

export type PickerMode =
	| 'libraryFirst' // pick Library -> (optional Folder) -> Content Type
	| 'contentTypeFirst' // pick Content Type -> resolve Library (or ask if multiple)
	| 'mixed'; // both, with tabs/toggle

export type OverwritePolicy = 'overwrite' | 'skip' | 'suffix'; // filename (1).ext

export type UploadSelectionScope = 'single' | 'multiple';

export type LauncherEventMode = RenderMode;

export type WebPartStage = 'idle' | 'destination' | 'uploading' | 'editing';

// Add ComponentStage type (used by hooks)
export type ComponentStage = 'idle' | 'destination' | 'upload' | 'editing';

// ---------- Web part configuration (Property Pane) ----------

/** A single configured library option */
export interface LibraryOption {
	/** Server-relative URL to the library root, e.g. "/sites/Contoso/Shared Documents" */
	serverRelativeUrl: string;
	/** Optional friendly label to show users (otherwise we fall back to library Title) */
	label?: string;
	/** Optional default subfolder relative to the library root, e.g. "Invoices/2025" */
	defaultFolder?: string;
	/** Optional minimal view id for bulk edit pane context (no braces) */
	minimalViewId?: string;
	/** Which content types are allowed in this library (empty => derive from library or use CT map) */
	allowedContentTypeIds?: string[] | 'all';
}

/** Content type metadata we cache for pickers */
export interface ContentTypeInfo {
	id: string; // full CT ID (e.g., "0x0101...")
	name: string; // "Invoice"
	description?: string; // optional description to show under the picker
	hidden?: boolean; // do not show if true
	sealed?: boolean; // read-only CTs
	group?: string; // e.g., "Custom Content Types"
}

/** Mapping from CT -> possible libraries, built at runtime */
export interface ContentTypeToLibrariesMap {
	[contentTypeId: string]: string[]; // array of library serverRelativeUrl that support this CT
}

/** Mapping from Library -> CTs, built at runtime (resolved set after filtering hidden/sealed if needed) */
export interface LibraryToContentTypesMap {
	[serverRelativeUrl: string]: ContentTypeInfo[];
}

/** Web part property bag (persisted via property pane) */
export interface UploadAndEditWebPartProps {
	// Mode & behavior
	pickerMode: PickerMode;
	renderMode: RenderMode;
	selectionScope: UploadSelectionScope; // single | multiple
	allowFolderSelection: boolean;
	showContentTypePicker: boolean; // can hide CT picker entirely
	overwritePolicy: OverwritePolicy;

	// Configurable libraries
	libraries: LibraryOption[]; // one or many; if one, we show as static text
	defaultLibrary?: string; // serverRelativeUrl (optional default)

	// If admins prefer CT-first UX, we can pre-limit CTs site-wide as a convenience
	globalAllowedContentTypeIds?: string[] | 'all';

	// Editor behavior
	enableBulkAutoRefresh: boolean;
	bulkWatchAllItems: boolean;

	// UI
	buttonLabel?: string; // "Upload files"
	dropzoneHint?: string; // "Drag & drop files here"
	successToast?: string; // "Uploaded! Opening properties…"

	// Accessibility & perf toggles
	disableDomNudges: boolean; // don't poke DOM to force pane/save
	sandboxExtra?: string; // extra sandbox flags for iframe
}

// ---------- Runtime selections (from pickers) ----------

/** The user's destination selection for a batch */
export interface DestinationChoice {
	libraryUrl: string;
	contentTypeId?: string;
	folderPath?: string;
	libraryTitle?: string;
	contentTypeName?: string;
}

/** A single file queued for upload (pre-flight) */
export interface PendingFile {
	file: File;
	/** Optional server name override (e.g., if you apply "(1)" suffix) */
	targetFileName?: string;
}

/** Per-file progress state (UI) */
export interface FileProgress {
	fileName: string;
	percent: number; // 0..100
	status: 'queued' | 'starting' | 'uploading' | 'done' | 'error' | 'skipped';
	errorMessage?: string;
	itemId?: number; // set when finished
}

/** Internal per-row state for UploadZone */
export interface FileUploadState {
	file: File;
	fileName: string;
	percent: number; // 0..100
	status: 'queued' | 'starting' | 'uploading' | 'done' | 'error' | 'skipped';
	errorMessage?: string;
	itemId?: number;
	attempts: number;
}

/** Batch upload outcome (emitted to parent) */
export interface UploadBatchResult {
	itemIds: number[];
	failed: Array<{ name: string; message: string }>;
	skipped?: string[];
}

// ---------- Hook-related types ----------

/** Component state for hooks */
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

/** Hook return type for useUploadState */
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

/** Component callbacks for hooks */
export interface ComponentCallbacks {
	onFilesPicked?: (files: File[]) => void;
	onUploadStart?: (files: File[]) => void;
	onUploadProgress?: (progress: FileProgress[]) => void;
	onUploadComplete?: (result: UploadBatchResult) => void;
	onEditingStart?: (itemIds: number[]) => void;
	onEditingComplete?: () => void;
	onError?: (error: SPFormError) => void;
}

/** Error handling interface */
export interface SPFormError {
	message: string;
	code?: string;
	context?: any;
	originalError?: any;
}

/** Configuration validation result */
export interface ConfigValidationResult {
	isValid: boolean;
	errors: string[];
	warnings: string[];
}

// ---------- Services contracts ----------

export interface SharePointService {
	getLibraryContentTypes(libraryUrl: string): Promise<ContentTypeInfo[]>;
	getLibraryTitle(libraryUrl: string): Promise<string>;

	uploadFileWithProgress(
		libraryUrl: string,
		folderPath: string | undefined,
		file: File,
		onProgress: (pct: number) => void,
		overwritePolicy: OverwritePolicy,
		chunkSizeBytes?: number,
		confirmOverwrite?: (fileName: string) => Promise<boolean>,
		contentTypeId?: string
	): Promise<{ itemId: number; serverRelativeUrl: string; uniqueId: string }>;

	setItemContentType(libraryUrl: string, itemId: number, contentTypeId: string): Promise<void>;

	fileExists(
		libraryUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean>;
}

// ---------- Launcher/Editor events ----------

export type LauncherDeterminedInfo = { mode: 'single' | 'bulk'; url: string; bulk: boolean };
export type LauncherOpenInfo = { mode: 'single' | 'bulk'; url: string };
export interface LauncherMetrics {
	msToDetermined?: number;
	msToOpen?: number;
}

// ---------- Component state shapes ----------

export interface UploadAndEditState {
	stage: WebPartStage;
	// destination picker
	choice?: DestinationChoice;
	// upload queue
	pending: PendingFile[];
	progress: FileProgress[];
	overallPct: number;
	// results
	uploadedItemIds: number[];
	// editor control
	showEditor: boolean;
}

// ---------- Utility helpers contracts (optional) ----------

export interface TelemetryTimer {
	start(label: string): void;
	end(label: string): number; // ms
}

export interface NormalizedError {
	message: string;
	cause?: unknown;
	code?: string | number;
}
