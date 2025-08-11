import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog, DialogType, IconButton, Spinner } from '@fluentui/react';
import { RenderMode } from '../../types';

export interface LibraryItemEditorLauncherProps {
	siteUrl: string;
	libraryServerRelativeUrl: string;
	itemIds: number[];
	contentTypeId?: string;
	viewId?: string; // used for bulk context if applicable
	renderMode: RenderMode; // 'modal' | 'samepage' | 'newtab'
	isOpen?: boolean; // modal only
	spfxContext: any;

	// ✅ Align with types.ts: { mode, url }
	onDetermined?: (info: { mode: RenderMode; url: string }) => void;
	onOpen?: (info: { mode: RenderMode; url: string }) => void;

	onSaved?: () => void;
	onDismiss?: () => void;

	// (hooks kept for future expansion—even if not used deeply here)
	enableBulkAutoRefresh?: boolean;
	bulkWatchAllItems?: boolean;
	disableDomNudges?: boolean;

	// iframe niceties
	sandboxExtra?: string;
	autoHeightBestEffort?: boolean;
}

export const LibraryItemEditorLauncher: React.FC<LibraryItemEditorLauncherProps> = ({
	siteUrl,
	libraryServerRelativeUrl,
	itemIds,
	contentTypeId,
	viewId,
	renderMode,
	isOpen = true,
	onDetermined,
	onOpen,
	onSaved,
	onDismiss,
	enableBulkAutoRefresh,
	bulkWatchAllItems,
	disableDomNudges,
	sandboxExtra,
	autoHeightBestEffort,
}) => {
	const [targetUrl, setTargetUrl] = useState<string | null>(null);
	const iframeRef = useRef<HTMLIFrameElement>(null);
	const isBulk = itemIds.length > 1;

	// Force re-layout when modal opens
	const vpTick = useMemo(() => Date.now(), [isOpen]);

	// Build the URL for single or bulk edit (modern experience)
	useEffect(() => {
		if (!itemIds?.length) return;

		let url: string;

		if (isBulk) {
			// Modern bulk edit by navigating to the library view with selected items.
			// NOTE: You may tailor this to your tenant/form setup; this works as a baseline.
			const base = `${siteUrl}${libraryServerRelativeUrl.replace(/\/$/, '')}/Forms/AllItems.aspx`;
			const u = new URL(base, window.location.origin);
			if (viewId) u.searchParams.set('viewid', viewId);
			// select param uses item ids; some tenants need the "id_.000" style—adjust if needed.
			u.searchParams.set('select', itemIds.join(','));
			u.searchParams.set('editAll', '1');
			url = u.toString();
		} else {
			// Single item modern edit form
			const u = new URL(`${siteUrl}/_layouts/15/listform.aspx`, window.location.origin);
			u.searchParams.set('PageType', '6'); // edit
			// ListId is not strictly required if we use ListUrl, but server-relative path works for most
			u.searchParams.set('ListUrl', libraryServerRelativeUrl);
			u.searchParams.set('Id', String(itemIds[0]));
			u.searchParams.set('action', 'edit');
			if (contentTypeId) u.searchParams.set('ContentTypeId', contentTypeId);
			url = u.toString();
		}

		// Add Source to bring user back (useful for samepage/newtab; harmless for modal)
		const currentUrl = window.location.href;
		const withSource = new URL(url, window.location.origin);
		withSource.searchParams.set('Source', currentUrl);

		const finalUrl = withSource.toString();
		setTargetUrl(finalUrl);

		// ✅ Fire aligned event
		onDetermined?.({ mode: renderMode, url: finalUrl });

		// Auto-navigate for non-modal modes
		if (renderMode === 'samepage') {
			onOpen?.({ mode: 'samepage', url: finalUrl });
			window.location.assign(finalUrl);
		} else if (renderMode === 'newtab') {
			onOpen?.({ mode: 'newtab', url: finalUrl });
			window.open(finalUrl, '_blank', 'noopener');
			onDismiss?.();
		}

		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [siteUrl, libraryServerRelativeUrl, itemIds, contentTypeId, viewId, renderMode]);

	const onIframeLoad = React.useCallback(() => {
		// ✅ Aligned event
		onOpen?.({ mode: 'modal', url: targetUrl || '' });

		if (autoHeightBestEffort) {
			try {
				iframeRef.current?.contentWindow?.focus();
			} catch {}
		}

		// If you want to detect save & close automatically, you'd need DOM inspection or Source redirect detection.
		// We keep that out here to stay tenant-safe/minimal.
	}, [onOpen, targetUrl, autoHeightBestEffort]);

	// Render modes
	if (renderMode !== 'modal') return null;

	return (
		<Dialog
			hidden={!isOpen}
			onDismiss={onDismiss}
			minWidth="80%"
			maxWidth="95%"
			dialogContentProps={{
				type: DialogType.close,
				title: isBulk ? 'Edit properties (bulk)' : 'Edit properties',
			}}
			modalProps={{ isBlocking: true }}
		>
			{!targetUrl && <Spinner label="Preparing editor..." />}
			{targetUrl && (
				<iframe
					key={vpTick}
					ref={iframeRef}
					title="Edit properties frame"
					src={targetUrl}
					onLoad={onIframeLoad}
					style={{
						width: '100%',
						height: autoHeightBestEffort ? 'clamp(70vh, 85vh, 95vh)' : '600px',
						border: 'none',
					}}
					sandbox={`allow-scripts allow-same-origin allow-forms allow-popups${
						sandboxExtra ? ` ${sandboxExtra}` : ''
					}`}
				/>
			)}
			<IconButton
				iconProps={{ iconName: 'Cancel' }}
				title="Close"
				ariaLabel="Close"
				onClick={onDismiss}
				style={{ position: 'absolute', top: 8, right: 8 }}
			/>
		</Dialog>
	);
};
