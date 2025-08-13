// components/UploadZone.tsx
import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import {
	Stack,
	Text,
	ProgressIndicator,
	MessageBar,
	MessageBarType,
	PrimaryButton,
	DefaultButton,
	Link,
	Icon,
} from '@fluentui/react';
import {
	DestinationChoice,
	SharePointService,
	OverwritePolicy,
	UploadBatchResult,
	FileProgress,
	FileUploadState,
} from '../types';

export interface UploadZoneProps {
	destination: DestinationChoice;
	spService: SharePointService;
	allowMultiple?: boolean;
	overwritePolicy?: OverwritePolicy;
	initialFiles?: File[];
	maxConcurrentUploads?: number;

	onBatchComplete?: (result: UploadBatchResult) => void;
	onBatchCanceled?: () => void;
	onProgress?: (progress: FileProgress[]) => void;

	confirmOverwrite?: (fileName: string) => Promise<boolean>;

	title?: string;
	hint?: string;
	cancelButtonLabel?: string;
	autoStart?: boolean;
}

type OneResult =
	| { status: 'done'; itemId: number }
	| { status: 'skipped' }
	| { status: 'error'; errorMessage: string };

export const UploadZone: React.FC<UploadZoneProps> = (props) => {
	const {
		destination,
		spService,
		allowMultiple = true,
		overwritePolicy = 'suffix',
		initialFiles = [],
		maxConcurrentUploads = 3,
		onBatchComplete,
		onBatchCanceled,
		onProgress,
		confirmOverwrite,
		title = 'Uploading files...',
		hint,
		cancelButtonLabel = 'Cancel',
		autoStart = true,
	} = props;

	const [files, setFiles] = useState<FileUploadState[]>([]);
	const [isUploading, setIsUploading] = useState(false);
	const [isCanceled, setIsCanceled] = useState(false);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	const activeUploadsRef = useRef<Set<string>>(new Set());
	const isMountedRef = useRef(true);
	const startingRef = useRef<boolean>(false); // prevent double start

	useEffect(() => {
		const next: FileUploadState[] = initialFiles.map((file) => ({
			file,
			fileName: file.name,
			percent: 0,
			status: 'queued',
			attempts: 0,
		}));
		setFiles(next);
		setIsUploading(false);
		setIsCanceled(false);
		setErrorMsg(null);
		startingRef.current = false;
	}, [initialFiles]);

	useEffect(() => {
		return () => {
			isMountedRef.current = false;
			activeUploadsRef.current.clear();
		};
	}, []);

	const overallProgress = useMemo(() => {
		if (files.length === 0) return 0;
		const nonSkipped = files.filter((f) => f.status !== 'skipped');
		if (!nonSkipped.length) return 100;
		const total = nonSkipped.reduce((sum, f) => sum + (isFinite(f.percent) ? f.percent : 0), 0);
		return Math.round(total / nonSkipped.length);
	}, [files]);

	const updateFileState = useCallback((index: number, updates: Partial<FileUploadState>) => {
		if (!isMountedRef.current) return;
		setFiles((prev) => {
			const arr = [...prev];
			if (arr[index]) arr[index] = { ...arr[index], ...updates };
			return arr;
		});
	}, []);

	const uploadOne = useCallback(
		async (index: number): Promise<OneResult> => {
			const snapshot = files[index];
			if (!snapshot || isCanceled || !isMountedRef.current) return { status: 'skipped' };

			const fileId = `${index}-${snapshot.file.name}`;
			activeUploadsRef.current.add(fileId);

			// show an indeterminate “Starting…” until the first progress tick
			updateFileState(index, { status: 'starting', percent: 0, errorMessage: undefined });

			// throttle progress updates to ~10fps
			let last = 0;
			const onProgressUpdate = (percent: number) => {
				const now = Date.now();
				if (now - last > 100) {
					last = now;
					updateFileState(index, {
						status: percent >= 0 && percent < 100 ? 'uploading' : 'uploading',
						percent: Math.max(0, Math.min(100, percent)),
					});
				}
			};

			try {
				const { itemId } = await spService.uploadFileWithProgress(
					destination.libraryUrl,
					destination.folderPath,
					snapshot.file,
					onProgressUpdate,
					overwritePolicy,
					2 * 1024 * 1024, // 2MB chunks => earlier callbacks
					confirmOverwrite,
					destination.contentTypeId // 👈 set CT immediately per-file
				);

				if (!isMountedRef.current) return { status: 'skipped' };

				updateFileState(index, {
					status: 'done',
					percent: 100,
					itemId,
					attempts: snapshot.attempts + 1,
				});
				return { status: 'done', itemId };
			} catch (e: any) {
				if (!isMountedRef.current) return { status: 'skipped' };

				// Prefer our service’s skip flag if present
				const isSkip =
					e && (e.__skip__ === true || /already exists|policy=skip/i.test(e.message || ''));
				if (isSkip) {
					updateFileState(index, {
						status: 'skipped',
						errorMessage: 'Skipped',
						attempts: snapshot.attempts + 1,
						percent: 0,
					});
					return { status: 'skipped' };
				}

				const message = (e && (e.message || e.toString())) || 'Upload failed';
				updateFileState(index, {
					status: 'error',
					errorMessage: message,
					attempts: snapshot.attempts + 1,
				});
				return { status: 'error', errorMessage: message };
			} finally {
				activeUploadsRef.current.delete(fileId);
			}
		},
		[
			files,
			isCanceled,
			destination.libraryUrl,
			destination.folderPath,
			destination.contentTypeId,
			overwritePolicy,
			confirmOverwrite,
			spService,
			updateFileState,
		]
	);

	const retryUpload = useCallback(
		async (index: number) => {
			if (isUploading || isCanceled) return;
			updateFileState(index, { status: 'queued', errorMessage: undefined, percent: 0 });
			await uploadOne(index);
		},
		[isUploading, isCanceled, updateFileState, uploadOne]
	);

	const startUpload = useCallback(async () => {
		if (files.length === 0 || isUploading || startingRef.current) return;
		startingRef.current = true;

		setIsUploading(true);
		setIsCanceled(false);
		setErrorMsg(null);

		// local result accumulator (don’t rely on React state timing)
		const itemIds: number[] = [];
		const failed: Array<{ name: string; message: string }> = [];
		const skipped: string[] = [];

		try {
			const targets = files
				.map((f, i) => ({ f, i }))
				.filter(({ f }) => f.status === 'queued' || f.status === 'error');

			let inFlight = 0;
			let cursor = 0;

			const launchNext = async (): Promise<void> => {
				if (isCanceled || !isMountedRef.current) return;
				if (cursor >= targets.length) return;

				const idx = targets[cursor++].i;
				inFlight++;

				const res = await uploadOne(idx);
				if (res.status === 'done') itemIds.push(res.itemId);
				else if (res.status === 'error')
					failed.push({ name: files[idx].file.name, message: res.errorMessage });
				else skipped.push(files[idx].file.name);

				inFlight--;
				if (cursor < targets.length) await launchNext();
			};

			const starters = Math.min(maxConcurrentUploads, targets.length);
			const runners: Promise<void>[] = [];
			for (let s = 0; s < starters; s++) runners.push(launchNext());
			await Promise.all(runners);
		} catch (e: any) {
			if (isMountedRef.current) setErrorMsg(e?.message || 'Batch upload failed');
		} finally {
			if (isMountedRef.current) {
				setIsUploading(false);
				startingRef.current = false;

				if (!isCanceled && onBatchComplete) {
					const result: UploadBatchResult = { itemIds, failed, skipped };
					onBatchComplete(result);
				}
			}
		}
	}, [files, isUploading, isCanceled, maxConcurrentUploads, uploadOne, onBatchComplete]);

	const cancelUpload = useCallback(() => {
		setIsCanceled(true);
		setIsUploading(false);
		activeUploadsRef.current.clear();
		onBatchCanceled?.();
	}, [onBatchCanceled]);

	useEffect(() => {
		if (autoStart && files.length > 0 && !isUploading && !isCanceled) {
			const hasQueued = files.some((f) => f.status === 'queued');
			if (hasQueued) startUpload();
		}
	}, [autoStart, files, isUploading, isCanceled, startUpload]);

	useEffect(() => {
		if (!onProgress) return;
		const progress: FileProgress[] = files.map((f) => ({
			fileName: f.fileName,
			percent: f.percent,
			status: f.status,
			errorMessage: f.errorMessage,
			itemId: f.itemId,
		}));
		onProgress(progress);
	}, [files, onProgress]);

	const getStatusIcon = (status: FileProgress['status']) => {
		switch (status) {
			case 'done':
				return <Icon iconName="CheckMark" style={{ color: '#107c10' }} />;
			case 'error':
				return <Icon iconName="ErrorBadge" style={{ color: '#d13438' }} />;
			case 'skipped':
				return <Icon iconName="Info" style={{ color: '#ffb900' }} />;
			case 'uploading':
			case 'starting':
				return <Icon iconName="Upload" style={{ color: '#0078d4' }} />;
			default:
				return <Icon iconName="Clock" style={{ color: '#605e5c' }} />;
		}
	};

	const formatFileSize = (bytes: number): string => {
		if (bytes === 0) return '0 Bytes';
		const k = 1024;
		const sizes = ['Bytes', 'KB', 'MB', 'GB'];
		const i = Math.floor(Math.log(bytes) / Math.log(k));
		return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
	};

	return (
		<Stack tokens={{ childrenGap: 16 }}>
			{/* Destination header */}
			<Stack
				styles={{
					root: {
						padding: 16,
						backgroundColor: '#f8f9fa',
						borderRadius: 8,
						border: '1px solid #edebe9',
					},
				}}
			>
				<Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
					{destination.libraryTitle || destination.libraryUrl}
				</Text>
				{destination.contentTypeName && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						Content type: <strong>{destination.contentTypeName}</strong>
					</Text>
				)}
				{destination.folderPath && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						Folder: <strong>{destination.folderPath}</strong>
					</Text>
				)}
			</Stack>

			{/* Header */}
			<Stack tokens={{ childrenGap: 8 }}>
				<Text variant="large" styles={{ root: { fontWeight: 600 } }}>
					{title}
				</Text>
				{hint && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						{hint}
					</Text>
				)}
			</Stack>

			{/* Error message */}
			{errorMsg && (
				<MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg(null)}>
					{errorMsg}
				</MessageBar>
			)}

			{/* Overall progress */}
			{files.length > 0 && (
				<Stack tokens={{ childrenGap: 8 }}>
					<Text variant="medium">
						Uploading {files.length} file{files.length > 1 ? 's' : ''}
					</Text>
					<ProgressIndicator
						percentComplete={overallProgress / 100}
						description={`${overallProgress}% complete`}
					/>
				</Stack>
			)}

			{/* Per-file rows */}
			<Stack tokens={{ childrenGap: 12 }}>
				{files.map((f, index) => (
					<Stack key={`${f.file.name}-${index}`} tokens={{ childrenGap: 4 }}>
						<Stack horizontal horizontalAlign="space-between" verticalAlign="center">
							<Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
								{getStatusIcon(f.status)}
								<Stack>
									<Text styles={{ root: { fontWeight: 500 } }}>{f.file.name}</Text>
									<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
										{formatFileSize(f.file.size)}
									</Text>
								</Stack>
							</Stack>

							{f.status === 'error' && (
								<Link
									onClick={() => retryUpload(index)}
									disabled={isUploading}
									styles={{ root: { fontWeight: 600 } }}
								>
									Retry
								</Link>
							)}
						</Stack>

						<ProgressIndicator
							percentComplete={
								f.status === 'starting' ? undefined : (isFinite(f.percent) ? f.percent : 0) / 100
							}
							description={
								f.status === 'error'
									? f.errorMessage || 'Failed'
									: f.status === 'done'
									? 'Completed'
									: f.status === 'skipped'
									? 'Skipped'
									: f.status === 'starting'
									? 'Starting…'
									: `${f.percent}%`
							}
						/>
					</Stack>
				))}
			</Stack>

			{/* Actions */}
			<Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
				{!autoStart && (
					<PrimaryButton
						text="Start Upload"
						onClick={startUpload}
						disabled={isUploading || files.length === 0}
						iconProps={{ iconName: 'Upload' }}
					/>
				)}
				<DefaultButton
					text={cancelButtonLabel}
					onClick={cancelUpload}
					disabled={!isUploading && files.length === 0}
				/>
			</Stack>
		</Stack>
	);
};
