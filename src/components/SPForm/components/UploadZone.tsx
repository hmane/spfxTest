import {
	DefaultButton,
	Link,
	MessageBar,
	MessageBarType,
	PrimaryButton,
	ProgressIndicator,
	Stack,
	Text,
} from '@fluentui/react';
import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DestinationChoice, OverwritePolicy, SharePointService, UploadBatchResult } from '../types';
import { debounce } from '../utils';
type OneResult =
	| { status: 'done'; itemId: number }
	| { status: 'skipped' }
	| { status: 'error'; errorMessage: string };
export interface UploadZoneProps {
	destination: DestinationChoice; // may include libraryTitle/contentTypeName/folderPath
	spService: SharePointService;
	allowMultiple?: boolean;
	overwritePolicy?: OverwritePolicy;
	initialFiles?: File[];

	onBatchComplete?: (result: UploadBatchResult) => void;
	onBatchCanceled?: () => void;

	/** Optional: if overwrite policy is "overwrite", ask user first */
	confirmOverwrite?: (fileName: string) => Promise<boolean>;

	title?: string;
	hint?: string;
	cancelButtonLabel?: string;

	/** Start uploading immediately when mounted */
	autoStart?: boolean; // default: true

	/**
	 * NEW: If false, do NOT call onBatchComplete automatically.
	 * Keep the UploadZone visible so users can retry failed items, then
	 * call onRequestProceed() when you want to move on.
	 */
	launchEditorOnComplete?: boolean; // default: true
	onRequestProceed?: (result: UploadBatchResult) => void;
}

type RowStatus = 'queued' | 'starting' | 'uploading' | 'done' | 'error' | 'skipped';

type FileRow = {
	file: File;
	targetFileName?: string;
	percent: number;
	status: RowStatus;
	errorMessage?: string;
	itemId?: number;
	attempts: number;
};

export const UploadZone: React.FC<UploadZoneProps> = (props) => {
	const {
		destination,
		spService,
		allowMultiple = true,
		overwritePolicy = 'suffix',
		initialFiles = [],
		onBatchComplete,
		onBatchCanceled,
		confirmOverwrite,
		title,
		hint,
		cancelButtonLabel = 'Cancel',
		autoStart = true,
		launchEditorOnComplete = true,
		onRequestProceed,
	} = props;

	const [rows, setRows] = useState<FileRow[]>(() =>
		initialFiles.map((f) => ({ file: f, percent: 0, status: 'queued', attempts: 0 }))
	);
	const [uploading, setUploading] = useState(false);
	const [canceled, setCanceled] = useState(false);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	// refresh when new files come in
	useEffect(() => {
		setRows(initialFiles.map((f) => ({ file: f, percent: 0, status: 'queued', attempts: 0 })));
		setUploading(false);
		setCanceled(false);
		setErrorMsg(null);
	}, [initialFiles]);

	// helper: compute batch summary
	const summary = useMemo(() => {
		const itemIds = rows
			.filter((r) => r.status === 'done' && typeof r.itemId === 'number')
			.map((r) => r.itemId!) as number[];
		const failed = rows
			.filter((r) => r.status === 'error')
			.map((r) => ({ name: r.file.name, message: r.errorMessage || 'Failed' }));
		const skipped = rows.filter((r) => r.status === 'skipped').map((r) => r.file.name);
		return { itemIds, failed, skipped };
	}, [rows]);

	const resetAndCancel = () => {
		setUploading(false);
		setCanceled(true);
		onBatchCanceled?.();
	};

	// add a return type for one-file attempt

	// CHANGE uploadOne: return OneResult instead of void
	const uploadOne = useCallback(
		async (index: number): Promise<OneResult> => {
			const row = rows[index];
			if (!row) return { status: 'skipped' };

			// mark starting
			setRows((prev) => {
				const n = [...prev];
				n[index] = { ...n[index], status: 'starting', percent: 0, errorMessage: undefined };
				return n;
			});

			// existence preflight (unchanged logic)
			try {
				const exists = await spService.fileExists(
					destination.libraryUrl,
					destination.folderPath,
					row.file.name
				);
				if (exists) {
					if (overwritePolicy === 'skip') {
						setRows((prev) => {
							const n = [...prev];
							n[index] = {
								...n[index],
								status: 'skipped',
								errorMessage: 'Already exists (skipped)',
								percent: 0,
							};
							return n;
						});
						return { status: 'skipped' };
					}
					if (overwritePolicy === 'overwrite' && confirmOverwrite) {
						const ok = await confirmOverwrite(row.file.name);
						if (!ok) {
							setRows((prev) => {
								const n = [...prev];
								n[index] = {
									...n[index],
									status: 'skipped',
									errorMessage: 'User chose not to overwrite',
									percent: 0,
								};
								return n;
							});
							return { status: 'skipped' };
						}
					}
					// suffix handled in service
				}
			} catch {
				// ignore and proceed
			}

			const updatePct = debounce((pct: number) => {
				setRows((prev) => {
					const n = [...prev];
					if (!n[index]) return n;
					n[index] = {
						...n[index],
						status: pct >= 0 && pct < 100 ? 'uploading' : n[index].status,
						percent: Math.max(0, Math.min(100, pct)),
					};
					return n;
				});
			}, 50);

			try {
				const { itemId } = await spService.uploadFileWithProgress(
					destination.libraryUrl,
					destination.folderPath,
					row.file,
					updatePct,
					overwritePolicy,
					2 * 1024 * 1024, // 2MB chunk for earlier callbacks
					overwritePolicy === 'overwrite' ? confirmOverwrite : undefined
				);

				setRows((prev) => {
					const n = [...prev];
					n[index] = {
						...n[index],
						status: 'done',
						percent: 100,
						itemId,
						attempts: row.attempts + 1,
					};
					return n;
				});
				return { status: 'done', itemId };
			} catch (e: any) {
				const message = (e && (e.message || e.toString())) || 'Upload failed';
				setRows((prev) => {
					const n = [...prev];
					n[index] = {
						...n[index],
						status: 'error',
						errorMessage: message,
						attempts: row.attempts + 1,
					};
					return n;
				});
				return { status: 'error', errorMessage: message };
			}
			// eslint-disable-next-line react-hooks/exhaustive-deps
		},
		[
			rows,
			destination.libraryUrl,
			destination.folderPath,
			destination.contentTypeId,
			overwritePolicy,
			confirmOverwrite,
			spService,
		]
	);

	// CHANGE startUpload: collect results locally and pass them up
	const startUpload = useCallback(async () => {
		if (!rows.length) return;
		setUploading(true);
		setCanceled(false);
		setErrorMsg(null);

		const itemIds: number[] = [];
		const failed: Array<{ name: string; message: string }> = [];
		const skipped: string[] = [];

		for (let i = 0; i < rows.length; i++) {
			if (canceled) break;
			const st = rows[i].status;
			if (st === 'queued' || st === 'starting' || st === 'error') {
				// eslint-disable-next-line no-await-in-loop
				const res = await uploadOne(i);
				if (res.status === 'done') {
					itemIds.push(res.itemId);
				} else if (res.status === 'error') {
					failed.push({ name: rows[i].file.name, message: res.errorMessage });
				} else {
					skipped.push(rows[i].file.name);
				}
			}
		}

		setUploading(false);
		if (canceled) return;

		const final = { itemIds, failed, skipped };

		if (launchEditorOnComplete) {
			onBatchComplete?.(final);
		} else {
			onRequestProceed?.(final);
		}
	}, [rows, canceled, uploadOne, launchEditorOnComplete, onBatchComplete, onRequestProceed]);

	// auto-start uploads when mounted / when rows change from picker
	const didAutoStartRef = useRef(false);
	useEffect(() => {
		if (!autoStart) return;
		if (rows.length > 0 && !uploading && !didAutoStartRef.current) {
			didAutoStartRef.current = true;
			startUpload();
		}
		if (rows.length === 0) {
			didAutoStartRef.current = false;
		}
	}, [rows, autoStart, uploading, startUpload]);

	// overall percent (only for files that are not skipped)
	const overallPct = useMemo(() => {
		const count = rows.filter((r) => r.status !== 'skipped').length;
		if (!count) return 0;
		const sum = rows
			.filter((r) => r.status !== 'skipped')
			.reduce((acc, r) => acc + (isFinite(r.percent) ? r.percent : 0), 0);
		return Math.round(sum / count);
	}, [rows]);

	const anyFailed = rows.some((r) => r.status === 'error');
	const anyDone = rows.some((r) => r.status === 'done');

	// ---- UI ----
	return (
		<Stack tokens={{ childrenGap: 12 }}>
			{/* Destination header */}
			<Stack
				styles={{
					root: {
						border: '1px solid #edebe9',
						borderRadius: 10,
						padding: 12,
						background: '#faf9f8',
					},
				}}
			>
				<Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
					{destination.libraryTitle || destination.libraryUrl}
				</Text>
				{destination.contentTypeName && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						Content type:&nbsp;
						<span style={{ fontWeight: 600 }}>{destination.contentTypeName}</span>
					</Text>
				)}
				{destination.folderPath && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						Folder:&nbsp;<span style={{ fontWeight: 600 }}>{destination.folderPath}</span>
					</Text>
				)}
			</Stack>

			<Text variant="large">{title || 'Uploading…'}</Text>
			{hint && (
				<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
					{hint}
				</Text>
			)}

			{errorMsg && (
				<MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg(null)}>
					{errorMsg}
				</MessageBar>
			)}

			{/* Overall progress (ignore skipped) */}
			{rows.length > 0 && (
				<ProgressIndicator
					label={`Uploading ${rows.length} file${rows.length > 1 ? 's' : ''}`}
					percentComplete={overallPct / 100}
				/>
			)}

			{/* Per-file rows */}
			<Stack tokens={{ childrenGap: 8 }}>
				{rows.map((r, idx) => (
					<Stack key={`${r.file.name}-${idx}`} tokens={{ childrenGap: 4 }}>
						<Stack horizontal horizontalAlign="space-between">
							<Text>{r.file.name}</Text>
							{r.status === 'error' && (
								<Link
									onClick={() => uploadOne(idx)}
									disabled={uploading}
									styles={{ root: { fontWeight: 600 } }}
								>
									Retry
								</Link>
							)}
						</Stack>

						<ProgressIndicator
							percentComplete={
								r.status === 'starting'
									? undefined // indeterminate while first chunk is getting underway
									: (r.percent || 0) / 100
							}
							description={
								r.status === 'error'
									? r.errorMessage || 'Failed'
									: r.status === 'done'
									? 'Completed'
									: r.status === 'skipped'
									? 'Skipped'
									: r.status === 'starting'
									? 'Starting…'
									: `${r.percent}%`
							}
						/>
					</Stack>
				))}
			</Stack>

			<Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
				{!launchEditorOnComplete && (
					<PrimaryButton
						text={anyDone ? 'Continue to properties' : 'Skip'}
						onClick={() => onRequestProceed?.(summary)}
						disabled={!anyDone && !anyFailed}
					/>
				)}
				<DefaultButton
					text={cancelButtonLabel}
					onClick={resetAndCancel}
					disabled={uploading && rows.length === 0}
				/>
			</Stack>
		</Stack>
	);
};
