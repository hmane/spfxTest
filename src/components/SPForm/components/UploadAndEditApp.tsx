import * as React from 'react';
import { useMemo, useState } from 'react';
import {
	Stack,
	Dialog,
	DialogType,
	MessageBar,
	MessageBarType,
	PrimaryButton,
	Text,
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
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DragDropFiles } from '@pnp/spfx-controls-react/lib/DragDropFiles';

type Props = {
	siteUrl: string;
	spfxContext: any;

	pickerMode: PickerMode;
	renderMode: RenderMode;
	selectionScope: UploadSelectionScope;
	showContentTypePicker: boolean;

	libraries: LibraryOption[];
	globalAllowedContentTypeIds?: string[] | 'all';

	overwritePolicy: OverwritePolicy;

	enableBulkAutoRefresh: boolean;
	bulkWatchAllItems: boolean;

	buttonLabel?: string;
	dropzoneHint?: string;
	successToast?: string;

	disableDomNudges: boolean;
	sandboxExtra?: string;

	showLoading?: (msg?: string) => void;
	hideLoading?: () => void;

	// Optional: async confirm overwrite hook (true = overwrite, false = skip)
	confirmOverwrite?: (fileName: string) => Promise<boolean>;
};

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

	const spService = useMemo(
		() => createSharePointService(),
		[siteUrl, spfxContext]
	);
	const { push } = useToasts();

	const [dialogOpen, setDialogOpen] = useState(false);
	const [stage, setStage] = useState<'destination' | 'upload' | 'editing' | 'idle'>('idle');
	const [choice, setChoice] = useState<DestinationChoice | undefined>();
	const [pendingFiles, setPendingFiles] = useState<File[]>([]);
	const [uploadedItemIds, setUploadedItemIds] = useState<number[]>([]);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	const isConfigured = useMemo(() => Array.isArray(libraries) && libraries.length > 0, [libraries]);

	// Button click uses a hidden input for “select files”
	const fileInputRef = React.useRef<HTMLInputElement>(null);
	const allowMultiple = selectionScope === 'multiple';

	const handleFilesPicked = (files: File[]) => {
		if (!files?.length) return;
		setPendingFiles(allowMultiple ? files : [files[0]]);
		setDialogOpen(true);
		setStage('destination');
	};

	// Destination
	const handleDestinationSubmit = (c: DestinationChoice) => {
		setChoice(c);
		setStage('upload');
	};
	const handleDestinationCancel = () => {
		setDialogOpen(false);
		setStage('idle');
		setPendingFiles([]);
		setUploadedItemIds([]);
	};

	// Upload
	const handleBatchComplete = (res: UploadBatchResult) => {
		setUploadedItemIds(res.itemIds);
		if (res.itemIds.length)
			push({
				kind: 'success',
				text: `Uploaded ${res.itemIds.length} file${
					res.itemIds.length > 1 ? 's' : ''
				}. Opening properties…`,
			});
		if (res.failed.length)
			push({
				kind: 'warning',
				text: `${res.failed.length} file${res.failed.length > 1 ? 's' : ''} skipped/failed.`,
			});
		showLoading?.('Preparing edit form…');
		setStage('editing');
	};
	const handleBatchCanceled = () => {
		setStage('destination');
	};

	// Editor
	const onLauncherDetermined = (_info: LauncherDeterminedInfo) => {
		hideLoading?.();
	};
	const onLauncherOpen = (_info: LauncherOpenInfo) => {};
	const onLauncherSaved = () => {
		push({ kind: 'success', text: (successToast && successToast.trim()) || 'Properties saved.' });
		setDialogOpen(false);
		setStage('idle');
		setPendingFiles([]);
		setUploadedItemIds([]);
	};
	const onLauncherDismiss = () => {
		setDialogOpen(false);
		setStage('idle');
		setPendingFiles([]);
		setUploadedItemIds([]);
		hideLoading?.();
	};

	// Inject per-library defaultFolder right before upload
	const choiceWithDefaultFolder = useMemo(() => {
		if (!choice) return undefined;
		const libCfg = libraries.find((l) => l.serverRelativeUrl === choice.libraryUrl);
		return { ...choice, folderPath: libCfg?.defaultFolder };
	}, [choice, libraries]);

	return (
		<Stack tokens={{ childrenGap: 16 }}>
			{!isConfigured ? (
				<Placeholder
					iconName="Edit"
					iconText="Configure this web part"
					description="Select one or more document libraries in the property pane."
					buttonLabel="Configure"
					onConfigure={() =>
						(window as any).SPPropertyPane && (window as any).SPPropertyPane.open()
					}
				/>
			) : (
				!dialogOpen && (
					<DragDropFiles onDrop={(files: Iterable<File> | ArrayLike<File>) => handleFilesPicked(Array.from(files))} dropEffect="copy">
						<Stack
							tokens={{ childrenGap: 8 }}
							styles={{
								root: {
									border: '1px dashed #c8c6c4',
									borderRadius: 10,
									padding: 16,
									background: '#fff',
									textAlign: 'center',
								},
							}}
						>
							<input
								ref={fileInputRef}
								type="file"
								multiple={allowMultiple}
								style={{ display: 'none' }}
								onChange={(e) => {
									const list = e.target.files;
									if (list?.length) handleFilesPicked(Array.from(list));
									e.currentTarget.value = '';
								}}
							/>
							<PrimaryButton
								text={buttonLabel || (allowMultiple ? 'Upload files' : 'Upload file')}
								onClick={() => fileInputRef.current?.click()}
							/>
							<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
								{dropzoneHint || (allowMultiple ? 'Drop files here' : 'Drop a file here')}
							</Text>
						</Stack>
					</DragDropFiles>
				)
			)}

			{/* One modal that hosts both steps */}
			<Dialog
				hidden={!dialogOpen}
				onDismiss={handleDestinationCancel}
				dialogContentProps={{ type: DialogType.close, title: undefined }}
				minWidth="50%"
				maxWidth="95%"
				modalProps={{ isBlocking: true }}
			>
				<Stack tokens={{ childrenGap: 16 }}>
					{errorMsg && (
						<MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg(null)}>
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
							subText="Choose the target library and content type."
						/>
					)}

					{stage === 'upload' && choiceWithDefaultFolder && (
						<UploadZone
							destination={choiceWithDefaultFolder}
							spService={spService}
							allowMultiple={allowMultiple}
							overwritePolicy={overwritePolicy}
							initialFiles={pendingFiles}
							onBatchComplete={handleBatchComplete}
							onBatchCanceled={handleBatchCanceled}
							title={buttonLabel}
							hint={dropzoneHint}
							confirmOverwrite={confirmOverwrite}
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
							onDetermined={onLauncherDetermined}
							onOpen={onLauncherOpen}
							onSaved={onLauncherSaved}
							onDismiss={onLauncherDismiss}
							enableBulkAutoRefresh={enableBulkAutoRefresh}
							bulkWatchAllItems={bulkWatchAllItems}
							disableDomNudges={disableDomNudges}
							sandboxExtra={sandboxExtra}
							autoHeightBestEffort
						/>
					)}
				</Stack>
			</Dialog>
		</Stack>
	);
};
