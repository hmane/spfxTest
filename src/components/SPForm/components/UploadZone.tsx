// src/webparts/UploadAndEdit/components/UploadZone.tsx
import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
	Stack,
	ProgressIndicator,
	Text,
	DefaultButton,
	MessageBar,
	MessageBarType,
} from '@fluentui/react';
import { debounce } from '../utils';
import { DestinationChoice, OverwritePolicy, SharePointService, UploadBatchResult } from '../types';

export interface UploadZoneProps {
	destination: DestinationChoice; // may include libraryTitle/contentTypeName/folderPath
	spService: SharePointService;
	allowMultiple?: boolean;
	overwritePolicy?: OverwritePolicy;
	initialFiles?: File[];

	onBatchComplete?: (result: UploadBatchResult) => void;
	onBatchCanceled?: () => void;

	confirmOverwrite?: (fileName: string) => Promise<boolean>;

	title?: string;
	hint?: string;
	cancelButtonLabel?: string;

	autoStart?: boolean; // default: true
}

type FileProgress = {
	name: string;
	percent: number;
	status: 'queued' | 'uploading' | 'done' | 'error';
	errorMessage?: string;
	itemId?: number;
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
	} = props;

	const [pending, setPending] = useState<{ file: File; targetFileName?: string }[]>(() =>
		initialFiles.map((f) => ({ file: f }))
	);
	const [progress, setProgress] = useState<FileProgress[]>([]);
	const [uploading, setUploading] = useState(false);
	const [canceled, setCanceled] = useState(false);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	useEffect(() => {
		setPending(initialFiles.map((f) => ({ file: f })));
		setProgress([]);
		setUploading(false);
		setCanceled(false);
		setErrorMsg(null);
	}, [initialFiles]);

	const resetBatch = () => {
		setUploading(false);
		setCanceled(false);
		setPending([]);
		setProgress([]);
		setErrorMsg(null);
	};

	const cancelUpload = () => {
		setCanceled(true);
		resetBatch();
		onBatchCanceled?.();
	};

	const startUpload = async () => {
		if (!destination?.libraryUrl || pending.length === 0) return;

		setUploading(true);
		setCanceled(false);
		setErrorMsg(null);

		const effectivePolicy: OverwritePolicy[] = pending.map(() => overwritePolicy);

		for (let i = 0; i < pending.length; i++) {
			const p = pending[i];
			try {
				const exists = await spService.fileExists(
					destination.libraryUrl,
					destination.folderPath,
					p.file.name
				);
				if (exists) {
					if (overwritePolicy === 'skip') {
						// keep skip
					} else if (overwritePolicy === 'suffix') {
						// keep suffix
					} else if (overwritePolicy === 'overwrite' && confirmOverwrite) {
						const ok = await confirmOverwrite(p.file.name);
						if (!ok) effectivePolicy[i] = 'skip';
					}
				}
			} catch {
				// ignore preflight errors
			}
			if (canceled) break;
		}

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
			const policy = effectivePolicy[i];

			try {
				setProgress((prev) => {
					const next = [...prev];
					next[i] = { ...next[i], status: 'uploading', percent: 0 };
					return next;
				});

				const updatePct = debounce((pct: number) => {
					setProgress((prev) => {
						const n = [...prev];
						n[i] = { ...n[i], percent: Math.max(0, Math.min(100, pct)), status: 'uploading' };
						return n;
					});
				}, 50);

				if (policy === 'skip') {
					const exists = await spService.fileExists(
						destination.libraryUrl,
						destination.folderPath,
						p.file.name
					);
					if (exists) {
						setProgress((prev) => {
							const n = [...prev];
							n[i] = { ...n[i], status: 'error', errorMessage: 'Skipped (name exists)' };
							return n;
						});
						failed.push({ name: p.file.name, message: 'Skipped (name exists)' });
						continue;
					}
				}

				const { itemId } = await spService.uploadFileWithProgress(
					destination.libraryUrl,
					destination.folderPath,
					p.file,
					(pct) => updatePct(pct),
					policy
				);

				setProgress((prev) => {
					const n = [...prev];
					n[i] = { ...n[i], percent: 100, status: 'done', itemId };
					return n;
				});
				itemIds.push(itemId);
			} catch (e: any) {
				const message = (e && (e.message || e.toString())) || 'Upload failed';
				failed.push({ name: p.file.name, message });
				setProgress((prev) => {
					const n = [...prev];
					n[i] = { ...n[i], status: 'error', errorMessage: message };
					return n;
				});
			}
		}

		setUploading(false);
		if (canceled) return;

		onBatchComplete?.({ itemIds, failed });

		resetBatch();
	};

	const didAutoStartRef = useRef(false);
	useEffect(() => {
		if (!autoStart) return;
		if (pending.length > 0 && !uploading && !didAutoStartRef.current) {
			didAutoStartRef.current = true;
			startUpload();
		}
		if (pending.length === 0) {
			didAutoStartRef.current = false;
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [pending, autoStart]);

	const overallPct = useMemo(() => {
		if (!progress.length) return 0;
		const sum = progress.reduce((acc, p) => acc + (isFinite(p.percent) ? p.percent : 0), 0);
		return Math.round(sum / progress.length);
	}, [progress]);

	const total = pending.length || progress.length;

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

			{total > 0 && (
				<ProgressIndicator
					label={`Uploading ${total} file${total > 1 ? 's' : ''}`}
					percentComplete={overallPct / 100}
				/>
			)}

			<Stack tokens={{ childrenGap: 6 }}>
				{progress.map((p, idx) => (
					<Stack key={`${p.name}-${idx}`} tokens={{ childrenGap: 2 }}>
						<Text>{p.name}</Text>
						<ProgressIndicator
							percentComplete={(p.percent || 0) / 100}
							description={
								p.status === 'error'
									? p.errorMessage || 'Failed'
									: p.status === 'done'
									? 'Completed'
									: `${p.percent}%`
							}
						/>
					</Stack>
				))}
			</Stack>

			<Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
				<DefaultButton
					text={cancelButtonLabel}
					onClick={cancelUpload}
					disabled={uploading && total === 0}
				/>
			</Stack>
		</Stack>
	);
};
