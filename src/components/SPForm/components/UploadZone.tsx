import {
	DefaultButton,
	Icon,
	MessageBar,
	MessageBarType,
	PrimaryButton,
	ProgressIndicator,
	Stack,
	Text,
} from '@fluentui/react';
import * as React from 'react';
import { useMemo, useRef, useState } from 'react';

import {
	DestinationChoice,
	FileProgress,
	OverwritePolicy,
	PendingFile,
	SharePointService,
	UploadBatchResult,
} from '../types';
import { debounce, formatBytes, normalizeError, overallPercent } from '../utils';
import styles from './UploadZone.module.scss';

export interface UploadZoneProps {
	/** Destination determined by the picker (required) */
	destination: DestinationChoice;

	/** SharePoint service bound to the current site (required) */
	spService: SharePointService;

	/** Allow selecting multiple files */
	allowMultiple?: boolean;

	/** Overwrite policy for collisions */
	overwritePolicy?: OverwritePolicy;

	/** Optional accept filter for input (e.g. ".docx,.pdf") */
	accept?: string;

	/** Initial files (e.g. passed from a prior UI) */
	initialFiles?: File[];

	/** Called when the entire batch completes (some may fail) */
	onBatchComplete?: (result: UploadBatchResult) => void;

	/** Called if user cancels the batch */
	onBatchCanceled?: () => void;

	confirmOverwrite?: (fileName: string) => Promise<boolean>;

	/** Optional: text overrides */
	title?: string; // e.g., "Upload files"
	hint?: string; // e.g., "Drag & drop files here, or click Select files"
	selectButtonLabel?: string;
	startButtonLabel?: string;
	cancelButtonLabel?: string;

	/** Optional: disable UI (e.g., while switching steps) */
	disabled?: boolean;
}

export const UploadZone: React.FC<UploadZoneProps> = (props) => {
	const {
		destination,
		spService,
		allowMultiple = true,
		overwritePolicy = 'suffix',
		accept,
		initialFiles = [],
		onBatchComplete,
		onBatchCanceled,

		title = 'Upload files',
		hint = 'Drag & drop here, or click Select files',
		selectButtonLabel = 'Select files',
		startButtonLabel = 'Start upload',
		cancelButtonLabel = 'Cancel',

		disabled = false,
	} = props;

	const [pending, setPending] = useState<PendingFile[]>(() =>
		initialFiles.map((f) => ({ file: f }))
	);
	const [progress, setProgress] = useState<FileProgress[]>([]);
	const [uploading, setUploading] = useState(false);
	const [canceled, setCanceled] = useState(false);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);
	const inputRef = useRef<HTMLInputElement>(null);

	const overallPct = useMemo(() => overallPercent(progress), [progress]);

	// ------------- File selection -------------

	const addFiles = (files: FileList | File[]) => {
		const arr = Array.from(files);
		setPending((prev) => {
			const next = allowMultiple ? [...prev] : [];
			for (const f of arr) {
				if (!allowMultiple && next.length >= 1) break;
				next.push({ file: f });
			}
			return next;
		});
		// Reset progress if we pick new files while not uploading
		if (!uploading) setProgress([]);
		setErrorMsg(null);
	};

	const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
		const list = e.target.files;
		if (list) addFiles(list);
		// keep input value reset so selecting same file again is possible
		e.currentTarget.value = '';
	};

	// drag & drop (simple)
	const onDrop = (e: React.DragEvent<HTMLDivElement>) => {
		e.preventDefault();
		if (disabled || uploading) return;
		if (e.dataTransfer?.files?.length) addFiles(e.dataTransfer.files);
	};
	const onDragOver = (e: React.DragEvent<HTMLDivElement>) => {
		e.preventDefault();
	};

	// ------------- Upload queue -------------

	const resetBatch = () => {
		setUploading(false);
		setCanceled(false);
	};

	const startUpload = async () => {
		if (!destination?.libraryUrl || pending.length === 0) return;

		setUploading(true);
		setCanceled(false);
		setErrorMsg(null);

		// ---- preflight duplicate names
		const effectivePolicyPerIndex: OverwritePolicy[] = pending.map(() => overwritePolicy);

		for (let i = 0; i < pending.length; i++) {
			const p = pending[i];
			const exists = await spService.fileExists(
				destination.libraryUrl,
				destination.folderPath,
				p.file.name
			);
			if (exists) {
				if (overwritePolicy === 'skip') {
					// keep skip
				} else if (overwritePolicy === 'suffix') {
					// keep suffix (no prompt)
				} else if (overwritePolicy === 'overwrite') {
					// optional prompt
					if (props.confirmOverwrite) {
						const ok = await props.confirmOverwrite(p.file.name);
						if (!ok) {
							// user declined → treat as skip
							effectivePolicyPerIndex[i] = 'skip';
						}
					}
				}
			}
			if (canceled) break;
		}

		// Seed progress rows
		const seeded: FileProgress[] = pending.map((p) => ({
			name: p.targetFileName || p.file.name,
			percent: 0,
			status: 'queued',
		}));
		setProgress(seeded);

		const failed: Array<{ name: string; message: string }> = [];
		const itemIds: number[] = [];

		for (let i = 0; i < pending.length; i++) {
			if (canceled) break;
			const p = pending[i];
			const policy = effectivePolicyPerIndex[i];

			try {
				setProgress((prev) => {
					const next = [...prev];
					next[i] = { ...next[i], status: 'uploading', percent: 0 };
					return next;
				});

				const updatePct = debounce((pct: number) => {
					setProgress((prev) => {
						const next = [...prev];
						next[i] = { ...next[i], percent: pct, status: 'uploading' };
						return next;
					});
				}, 50);

				// If skip on conflict and still exists, short-circuit
				if (
					policy === 'skip' &&
					(await spService.fileExists(destination.libraryUrl, destination.folderPath, p.file.name))
				) {
					setProgress((prev) => {
						const next = [...prev];
						next[i] = { ...next[i], status: 'error', errorMessage: 'Skipped (name exists)' };
						return next;
					});
					failed.push({ name: p.file.name, message: 'Skipped (name exists)' });
					continue;
				}

				const { itemId } = await spService.uploadFileWithProgress(
					destination.libraryUrl,
					destination.folderPath,
					p.file,
					(pct) => updatePct(Math.max(0, Math.min(100, pct))),
					policy
				);

				// Optional single-file CT force
				if (destination.contentTypeId && pending.length === 1) {
					try {
						await spService.setItemContentType(
							destination.libraryUrl,
							itemId,
							destination.contentTypeId
						);
					} catch {}
				}

				setProgress((prev) => {
					const next = [...prev];
					next[i] = { ...next[i], percent: 100, status: 'done', itemId };
					return next;
				});
				itemIds.push(itemId);
			} catch (e) {
				const msg = normalizeError(e).message;
				failed.push({ name: p.file.name, message: msg });
				setProgress((prev) => {
					const next = [...prev];
					next[i] = { ...next[i], status: 'error', errorMessage: msg };
					return next;
				});
			}
		}

		setUploading(false);
		if (canceled) {
			onBatchCanceled?.();
			return;
		}
		onBatchComplete?.({ itemIds, failed });
	};

	const cancelUpload = () => {
		// We can’t truly abort in-flight addChunked via PnPjs; this stops the queue after current file
		setCanceled(true);
	};

	// ------------- Render -------------

	const hasFiles = pending.length > 0;

	return (
		<Stack tokens={{ childrenGap: 16 }}>
			<Stack>
				<Text variant="xLarge">{title}</Text>
				<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
					{hint}
				</Text>
			</Stack>

			{errorMsg && (
				<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
					{errorMsg}
				</MessageBar>
			)}

			{/* Drop zone */}
			<Stack className={styles.zone} onDrop={onDrop} onDragOver={onDragOver}>
				<Stack horizontal horizontalAlign="space-between" verticalAlign="center">
					<Stack>
						<Text>Destination</Text>
						<Text variant="small">{destination.libraryUrl}</Text>
						{destination.folderPath && <Text variant="small">/{destination.folderPath}</Text>}
						{destination.contentTypeId && (
							<Text variant="small">CT: {destination.contentTypeId}</Text>
						)}
					</Stack>

					<Stack horizontal tokens={{ childrenGap: 8 }}>
						<input
							ref={inputRef}
							type="file"
							multiple={allowMultiple}
							accept={accept}
							onChange={onInputChange}
							style={{ display: 'none' }}
							disabled={disabled || uploading}
						/>
						<DefaultButton
							text={selectButtonLabel}
							onClick={() => inputRef.current?.click()}
							disabled={disabled || uploading}
						/>
						<PrimaryButton
							text={startButtonLabel}
							onClick={startUpload}
							disabled={disabled || uploading || !hasFiles}
						/>
						{uploading && <DefaultButton text={cancelButtonLabel} onClick={cancelUpload} />}
					</Stack>
				</Stack>
			</Stack>

			{/* Files list + per-file progress */}
			{hasFiles && (
				<Stack tokens={{ childrenGap: 8 }}>
					{pending.map((p, idx) => {
						const row = progress[idx];
						const pct = row?.percent ?? 0;
						const status = row?.status || 'queued';
						const name = p.file.name;
						const size = formatBytes(p.file.size);

						const iconName =
							status === 'done'
								? 'CompletedSolid'
								: status === 'error'
								? 'StatusErrorFull'
								: status === 'uploading'
								? 'Upload'
								: 'Edit';

						const iconColor =
							status === 'done' ? 'green' : status === 'error' ? '#a4262c' : '#605e5c';

						return (
							<Stack
								key={`${name}-${idx}`}
								className={styles.fileRow}
								horizontal
								verticalAlign="center"
							>
								<Icon iconName={iconName} styles={{ root: { color: iconColor } }} />
								<Stack grow>
									<Text>
										{name} <Text variant="small">({size})</Text>
									</Text>
									<ProgressIndicator percentComplete={pct / 100} />
								</Stack>
								<Stack styles={{ root: { width: 80, textAlign: 'right' } }}>
									<Text variant="small">{pct}%</Text>
								</Stack>
							</Stack>
						);
					})}
				</Stack>
			)}

			{/* Overall progress */}
			{uploading && progress.length > 0 && (
				<Stack>
					<Text>Overall progress</Text>
					<ProgressIndicator percentComplete={overallPct / 100} />
				</Stack>
			)}
		</Stack>
	);
};
