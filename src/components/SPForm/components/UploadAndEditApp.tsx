import * as React from 'react';
import { useCallback, useMemo, useRef, useState } from 'react';
import { useToasts } from './ToastHost';
import {
	Stack,
	Text,
	Dialog,
	DialogType,
	Spinner,
	MessageBar,
	MessageBarType,
} from '@fluentui/react';

import {
	ContentTypeInfo,
	DestinationChoice,
	FileProgress,
	LauncherMetrics,
	LauncherOpenInfo,
	LauncherDeterminedInfo,
	LibraryOption,
	OverwritePolicy,
	PickerMode,
	RenderMode,
	UploadAndEditState,
	UploadAndEditWebPartProps,
	UploadBatchResult,
	UploadSelectionScope,
} from '../types';

import { createSharePointService } from '../services/sharepoint';
import { DestinationPicker } from './DestinationPicker';
import { UploadZone } from './UploadZone';
import { LibraryItemEditorLauncher } from './editor/LibraryItemEditorLauncher';

type Props = {
	// SPFx context + site
	siteUrl: string;
	spfxContext: any;

	// Web part configuration
	pickerMode: PickerMode;
	renderMode: RenderMode;
	selectionScope: UploadSelectionScope;
	allowFolderSelection: boolean;
	showContentTypePicker: boolean;
	overwritePolicy: OverwritePolicy;

	libraries: LibraryOption[];
	defaultLibrary?: string;
	globalAllowedContentTypeIds?: string[] | 'all';

	// Editor behavior
	enableBulkAutoRefresh: boolean;
	bulkWatchAllItems: boolean;

	// UI bits
	buttonLabel?: string;
	dropzoneHint?: string;
	successToast?: string;

	// Accessibility & perf
	disableDomNudges: boolean;
	sandboxExtra?: string;

	// Minimal view for bulk edit if you want (per library overrides exist on LibraryOption too)
	minimalViewId?: string;

	// External loader hooks (optional)
	showLoading?: (msg?: string) => void;
	hideLoading?: () => void;
};

export const UploadAndEditApp: React.FC<Props> = (props) => {
	const {
		siteUrl,
		spfxContext,

		pickerMode,
		renderMode,
		selectionScope,
		allowFolderSelection,
		showContentTypePicker,
		overwritePolicy,

		libraries,
		defaultLibrary,
		globalAllowedContentTypeIds,

		enableBulkAutoRefresh,
		bulkWatchAllItems,

		buttonLabel = 'Upload files',
		dropzoneHint = 'Drag & drop files here, or click Select files',
		successToast,

		disableDomNudges,
		sandboxExtra,
		minimalViewId,

		showLoading,
		hideLoading,
	} = props;

    const { push } = useToasts();

	const spService = useMemo(
		() => createSharePointService(siteUrl, spfxContext),
		[siteUrl, spfxContext]
	);

	// -------------- Orchestration state --------------
	const [stage, setStage] = useState<'destination' | 'upload' | 'editing' | 'idle'>('destination');
	const [dialogOpen, setDialogOpen] = useState<boolean>(true);

	const [choice, setChoice] = useState<DestinationChoice | undefined>();
	const [uploadedItemIds, setUploadedItemIds] = useState<number[]>([]);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	// -------------- Stage 1: Destination --------------
	const handleDestinationSubmit = (c: DestinationChoice) => {
		setChoice(c);
		setStage('upload');
	};

	const handleDestinationCancel = () => {
		setDialogOpen(false);
		setStage('idle');
	};

	// -------------- Stage 1b: Upload --------------
	const handleBatchComplete = (res: UploadBatchResult) => {
		setUploadedItemIds(res.itemIds);
		if (res.itemIds.length)
			push({
				kind: 'success',
				text: `Uploaded ${res.itemIds.length} file(s). Opening properties…`,
			});
		if (res.failed.length)
			push({ kind: 'warning', text: `${res.failed.length} file(s) skipped/failed.` });
		setStage('editing');
	};



	const handleBatchCanceled = () => {
		// back to destination or close?
		setStage('destination');
	};

	// -------------- Stage 2: Edit properties --------------
	const onLauncherDetermined = (info: LauncherDeterminedInfo) => {
		// You wanted to hide loading here (component has computed where to go)
		hideLoading?.();
	};

	const onLauncherOpen = (info: LauncherOpenInfo) => {
		// Telemetry hook; usually no loader actions here
	};

	const onLauncherSaved = () => {
		push({ kind: 'success', text: 'Properties saved.' });
		setDialogOpen(false);
		setStage('idle');
	};

	const onLauncherDismiss = () => {
		setDialogOpen(false);
		setStage('idle');
	};

	// -------------- Derived: which minimal view to use --------------
	const viewIdForChoice = useMemo(() => {
		if (!choice?.libraryUrl) return undefined;
		const libCfg = libraries.find((l) => l.serverRelativeUrl === choice.libraryUrl);
		return libCfg?.minimalViewId || minimalViewId;
	}, [choice?.libraryUrl, libraries, minimalViewId]);

	// -------------- Render --------------
	return (
		<Dialog
			hidden={!dialogOpen}
			onDismiss={() => {
				setDialogOpen(false);
				setStage('idle');
			}}
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
						selectionScope={selectionScope}
						libraries={libraries}
						defaultLibrary={defaultLibrary}
						showContentTypePicker={showContentTypePicker}
						allowFolderSelection={allowFolderSelection}
						globalAllowedContentTypeIds={globalAllowedContentTypeIds}
						spService={spService}
						onSubmit={handleDestinationSubmit}
						onCancel={handleDestinationCancel}
						primaryText="Continue"
						cancelText="Cancel"
						title="Select destination"
						subText="Choose where your file(s) will be stored and (optionally) the content type."
					/>
				)}

				{stage === 'upload' && choice && (
					<UploadZone
						destination={choice}
						spService={spService}
						allowMultiple={selectionScope === 'multiple'}
						overwritePolicy={overwritePolicy}
						onBatchComplete={handleBatchComplete}
						onBatchCanceled={handleBatchCanceled}
						title={buttonLabel}
						hint={dropzoneHint}
						confirmOverwrite={async (fileName) => {
							// your own modal/confirm; return true to overwrite, false to skip
							//return await myConfirm(`"${fileName}" already exists. Replace it?`);
                            return true;
						}}
					/>
				)}

				{stage === 'editing' && choice && uploadedItemIds.length > 0 && (
					<LibraryItemEditorLauncher
						siteUrl={siteUrl}
						libraryServerRelativeUrl={choice.libraryUrl}
						itemIds={uploadedItemIds}
						contentTypeId={choice.contentTypeId}
						viewId={uploadedItemIds.length > 1 ? viewIdForChoice : undefined}
						renderMode={renderMode} // modal | samepage | newtab
						isOpen={renderMode === 'modal'} // only render modal content when needed
						spfxContext={spfxContext}
						// lifecycle
						onDetermined={onLauncherDetermined}
						onOpen={onLauncherOpen}
						onSaved={onLauncherSaved}
						onDismiss={onLauncherDismiss}
						// editor behavior
						enableBulkAutoRefresh={enableBulkAutoRefresh}
						bulkWatchAllItems={bulkWatchAllItems}
						disableDomNudges={disableDomNudges}
						sandboxExtra={sandboxExtra}
						// responsive niceties (good defaults baked in the launcher)
						autoHeightBestEffort
					/>
				)}
			</Stack>
		</Dialog>
	);
};
