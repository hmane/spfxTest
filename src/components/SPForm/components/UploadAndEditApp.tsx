// components/UploadAndEditApp.tsx - Fixed version
import * as React from 'react';
import { useMemo, useState, useCallback } from 'react';
import {
	Stack,
	Dialog,
	DialogType,
	MessageBar,
	MessageBarType,
	PrimaryButton,
	Text,
	Spinner,
	SpinnerSize,
} from '@fluentui/react';

import {
	DestinationChoice,
	LauncherDeterminedInfo,
	LauncherOpenInfo,
	LibraryOption,
	OverwritePolicy,
	PickerMode,
	RenderMode,
	UploadBatchResult,
	UploadSelectionScope,
} from '../types';

import { createSharePointService } from '../services/sharepoint';
import { DestinationPicker } from './DestinationPicker';
import { UploadZone } from './UploadZone';
import { LibraryItemEditorLauncher } from './editor/LibraryItemEditorLauncher';
import { useToasts } from './ToastHost';
import { useUploadState, useFileValidation, useErrorHandler } from '../hooks/useUploadState';
import { DragDropFiles } from '@pnp/spfx-controls-react/lib/DragDropFiles';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';

interface Props {
	siteUrl: string;
	spfxContext: any;

	// Component configuration
	pickerMode: PickerMode;
	renderMode: RenderMode;
	selectionScope: UploadSelectionScope;
	showContentTypePicker: boolean;

	// Library configuration
	libraries: LibraryOption[];
	globalAllowedContentTypeIds?: string[] | 'all';

	// Upload behavior
	overwritePolicy: OverwritePolicy;

	// Editor configuration
	enableBulkAutoRefresh: boolean;
	bulkWatchAllItems: boolean;

	// UI customization
	buttonLabel?: string;
	dropzoneHint?: string;
	successToast?: string;

	// Advanced options
	disableDomNudges: boolean;
	sandboxExtra?: string;

	// Optional callbacks
	showLoading?: (msg?: string) => void;
	hideLoading?: () => void;
	confirmOverwrite?: (fileName: string) => Promise<boolean>;
}

type AppStage = 'idle' | 'destination' | 'upload' | 'editing';

export const UploadAndEditApp: React.FC<Props> = (props) => {
	const {
		siteUrl,
		spfxContext,
		pickerMode,
		renderMode,
		selectionScope,
		showContentTypePicker,
		libraries,
		globalAllowedContentTypeIds,
		overwritePolicy,
		enableBulkAutoRefresh,
		bulkWatchAllItems,
		buttonLabel = 'Upload files',
		dropzoneHint = 'Drag & drop files here, or click Select files',
		successToast,
		disableDomNudges,
		sandboxExtra,
		showLoading,
		hideLoading,
		confirmOverwrite,
	} = props;

	// Memoize SharePoint service to prevent recreating on every render
	const spService = useMemo(() => {
		if (!spfxContext) {
			throw new Error('SPFx context is required');
		}
		return createSharePointService(spfxContext);
	}, [spfxContext]);

	const { push: pushToast } = useToasts();

	// Initialize hooks
	const { handleError } = useErrorHandler();

	const { validateFiles } = useFileValidation({
		maxFileSize: 250 * 1024 * 1024, // 250MB default limit
		maxFiles: selectionScope === 'single' ? 1 : 100,
	});

	const { state, actions } = useUploadState({
		onError: (error) => {
			pushToast({ kind: 'error', text: error.message });
		},
		onFilesPicked: (files) => {
			console.log(`${files.length} files selected for upload`);
		},
	});

	// Component state destructuring for easier access
	const {
		stage,
		dialogOpen,
		choice,
		pendingFiles,
		uploadedItemIds,
		errorMsg,
		isLoading,
		loadingMessage,
	} = state;

	const {
		setStage,
		setDialogOpen,
		setChoice,
		setPendingFiles,
		setUploadedItemIds,
		setError,
		setLoading,
		reset,
	} = actions;

	// Derived state
	const isConfigured = useMemo(() => Array.isArray(libraries) && libraries.length > 0, [libraries]);

	const allowMultiple = selectionScope === 'multiple';
	const fileInputRef = React.useRef<HTMLInputElement>(null);

	// Event handlers
	const handleFilesPicked = useCallback(
		(files: File[]) => {
			if (!files?.length) return;

			// Validate files using the validation hook
			const { valid, invalid } = validateFiles(files);

			// Show warnings for invalid files
			if (invalid.length > 0) {
				invalid.forEach(({ file, reason }) => {
					pushToast({
						kind: 'warning',
						text: `${file.name}: ${reason}`,
					});
				});
			}

			if (valid.length === 0) {
				pushToast({ kind: 'error', text: 'No valid files to upload' });
				return;
			}

			// Update state using the hook actions
			setPendingFiles(valid);
			setDialogOpen(true);
			setStage('destination');
			setError(null);
		},
		[validateFiles, pushToast, setPendingFiles, setDialogOpen, setStage, setError]
	);

	const handleDestinationSubmit = useCallback(
		(c: DestinationChoice) => {
			setChoice(c);
			setStage('upload');
		},
		[setChoice, setStage]
	);

	const handleDestinationCancel = useCallback(() => {
		reset(); // Use the reset action from the hook
	}, [reset]);

	const handleBatchComplete = useCallback(
		async (res: UploadBatchResult) => {
			try {
				setUploadedItemIds(res.itemIds);

				// Show upload results
				if (res.itemIds.length) {
					pushToast({
						kind: 'success',
						text: `Uploaded ${res.itemIds.length} file${res.itemIds.length > 1 ? 's' : ''}`,
					});
				}

				if (res.skipped?.length) {
					pushToast({
						kind: 'info',
						text: `${res.skipped.length} file${
							res.skipped.length > 1 ? 's' : ''
						} skipped (already existed)`,
					});
				}

				if (res.failed.length) {
					pushToast({
						kind: 'warning',
						text: `${res.failed.length} file${res.failed.length > 1 ? 's' : ''} failed to upload`,
					});
				}

				if (res.itemIds.length > 0) {
					setLoading(true, 'Preparing edit form...');
					setStage('editing');
				} else {
					// No files uploaded successfully
					reset();
				}
			} catch (error) {
				const formattedError = handleError(error, 'Upload Complete Handler');
				setError(formattedError.message);
			}
		},
		[
			choice,
			spService,
			pushToast,
			setUploadedItemIds,
			setLoading,
			setStage,
			reset,
			setError,
			handleError,
		]
	);

	const handleBatchCanceled = useCallback(() => {
		setStage('destination');
	}, [setStage]);

	// Editor event handlers
	const handleLauncherDetermined = useCallback(
		(_info: LauncherDeterminedInfo) => {
			setLoading(false);
		},
		[setLoading]
	);

	const handleLauncherOpen = useCallback((_info: LauncherOpenInfo) => {
		// Optional: track editor open event
	}, []);

	const handleLauncherSaved = useCallback(() => {
		const message = successToast?.trim() || 'Properties saved successfully';
		pushToast({ kind: 'success', text: message });
		reset();
	}, [successToast, pushToast, reset]);

	const handleLauncherDismiss = useCallback(() => {
		reset();
	}, [reset]);

	// Enrich choice with library defaults
	const enrichedChoice = useMemo(() => {
		if (!choice) return undefined;

		const libConfig = libraries.find((l) => l.serverRelativeUrl === choice.libraryUrl);
		return {
			...choice,
			folderPath: libConfig?.defaultFolder,
			libraryTitle: choice.libraryTitle || libConfig?.label,
		};
	}, [choice, libraries]);

	// Error boundary
	if (!spfxContext) {
		return (
			<MessageBar messageBarType={MessageBarType.error}>
				SPFx context is required but not provided
			</MessageBar>
		);
	}

	return (
		<Stack tokens={{ childrenGap: 16 }}>
			{!isConfigured ? (
				<Placeholder
					iconName="Edit"
					iconText="Configure this web part"
					description="Select one or more document libraries in the property pane."
					buttonLabel="Configure"
					onConfigure={() => {
						// Safely access SPPropertyPane
						const propertyPane = (window as any).SPPropertyPane;
						if (propertyPane?.open) {
							propertyPane.open();
						} else {
							pushToast({
								kind: 'warning',
								text: 'Property pane not available. Please edit the web part to configure.',
							});
						}
					}}
				/>
			) : (
				!dialogOpen && (
					<DragDropFiles
						onDrop={(files: File[] | FileList) =>
							handleFilesPicked(Array.isArray(files) ? files : Array.from(files as any))
						}
						dropEffect="copy"
					>
						<Stack
							tokens={{ childrenGap: 8 }}
							styles={{
								root: {
									border: '1px dashed #c8c6c4',
									borderRadius: 10,
									padding: 16,
									background: '#fff',
									textAlign: 'center',
									cursor: 'pointer',
									transition: 'border-color 0.2s ease',
									':hover': {
										borderColor: '#8a8886',
										background: '#faf9f8',
									},
								},
							}}
						>
							<input
								ref={fileInputRef}
								type="file"
								multiple={allowMultiple}
								style={{ display: 'none' }}
								onChange={(e) => {
									const fileList = e.target.files;
									if (fileList?.length) {
										handleFilesPicked(Array.from(fileList));
									}
									e.currentTarget.value = ''; // Reset for next selection
								}}
								accept="*/*" // Could be made configurable
							/>
							<PrimaryButton
								text={buttonLabel}
								onClick={() => fileInputRef.current?.click()}
								iconProps={{ iconName: 'Upload' }}
							/>
							<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
								{dropzoneHint}
							</Text>
						</Stack>
					</DragDropFiles>
				)
			)}

			<Dialog
				hidden={!dialogOpen}
				onDismiss={handleDestinationCancel}
				dialogContentProps={{
					type: DialogType.close,
					title: undefined,
					showCloseButton: true,
				}}
				minWidth="60%"
				maxWidth="95%"
				modalProps={{
					isBlocking: true,
					dragOptions: undefined, // Disable dragging for better UX
				}}
			>
				<Stack tokens={{ childrenGap: 16 }}>
					{/* Loading indicator */}
					{isLoading && (
						<Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
							<Spinner size={SpinnerSize.small} />
							<Text>{loadingMessage || 'Loading...'}</Text>
						</Stack>
					)}

					{errorMsg && (
						<MessageBar
							messageBarType={MessageBarType.error}
							onDismiss={() => setError(null)}
							isMultiline={false}
						>
							{errorMsg}
						</MessageBar>
					)}

					{stage === 'destination' && (
						<DestinationPicker
							pickerMode={pickerMode}
							libraries={libraries}
							showContentTypePicker={showContentTypePicker}
							globalAllowedContentTypeIds={globalAllowedContentTypeIds}
							spService={spService}
							onSubmit={handleDestinationSubmit}
							onCancel={handleDestinationCancel}
							primaryText="Continue"
							cancelText="Cancel"
							title="Select destination"
							subText={`Choose where to upload ${pendingFiles.length} file${
								pendingFiles.length > 1 ? 's' : ''
							}`}
						/>
					)}

					{stage === 'upload' && enrichedChoice && (
						<UploadZone
							destination={enrichedChoice}
							spService={spService}
							allowMultiple={allowMultiple}
							overwritePolicy={overwritePolicy}
							initialFiles={pendingFiles}
							onBatchComplete={handleBatchComplete}
							onBatchCanceled={handleBatchCanceled}
							title="Uploading files..."
							hint="Please wait while files are uploaded"
							confirmOverwrite={confirmOverwrite}
							autoStart={true}
						/>
					)}

					{stage === 'editing' && choice && uploadedItemIds.length > 0 && (
						<LibraryItemEditorLauncher
							siteUrl={siteUrl}
							libraryServerRelativeUrl={choice.libraryUrl}
							itemIds={uploadedItemIds}
							contentTypeId={choice.contentTypeId}
							renderMode={renderMode}
							isOpen={renderMode === 'modal'}
							spfxContext={spfxContext}
							onDetermined={handleLauncherDetermined}
							onOpen={handleLauncherOpen}
							onSaved={handleLauncherSaved}
							onDismiss={handleLauncherDismiss}
							enableBulkAutoRefresh={enableBulkAutoRefresh}
							bulkWatchAllItems={bulkWatchAllItems}
							disableDomNudges={disableDomNudges}
							sandboxExtra={sandboxExtra}
							autoHeightBestEffort={true}
							hideBreadcrumbs={true}
							hideContentTypeField={true}
						/>
					)}
				</Stack>
			</Dialog>
		</Stack>
	);
};
