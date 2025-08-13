// src/webparts/UploadAndEdit/components/editor/LibraryItemEditorLauncher.tsx

import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog, DialogType, IconButton, Stack } from '@fluentui/react';

import { spfi, SPFI } from '@pnp/sp';
import { SPFx as PnP_SPFX } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export type RenderMode = 'modal' | 'samepage' | 'newtab';

// If you already export these from types.ts, remove these and import instead.
export interface LauncherDeterminedInfo {
	mode: 'single' | 'bulk';
	url: string;
	bulk: boolean;
}
export interface LauncherOpenInfo {
	mode: 'single' | 'bulk';
	url: string;
}

export interface LibraryItemEditorLauncherProps {
	// Site + context
	siteUrl: string;
	spfxContext: any;

	// Target
	libraryServerRelativeUrl: string; // e.g. "/sites/Contoso/Shared Documents"
	itemIds: number[];
	contentTypeId?: string; // not used here (set during upload)

	// Optional: minimal view (GUID with no braces) for bulk
	viewId?: string;

	// Render behavior
	renderMode: RenderMode;
	isOpen?: boolean; // modal only

	// Lifecycle
	onDetermined?: (info: LauncherDeterminedInfo) => void;
	onOpen?: (info: LauncherOpenInfo) => void;
	onSaved?: () => void;
	onDismiss?: () => void;

	// Bulk niceties
	enableBulkAutoRefresh?: boolean;
	bulkWatchAllItems?: boolean;

	// DOM nudge controls
	disableDomNudges?: boolean;

	// Iframe sandbox extras
	sandboxExtra?: string;

	// Responsiveness
	autoHeightBestEffort?: boolean;
}

/** Resolve list GUID from server-relative library URL */
async function resolveListId(sp: SPFI, libraryServerRelativeUrl: string): Promise<string> {
	const list = await sp.web.getList(libraryServerRelativeUrl).select('Id')();
	return list?.Id as string;
}

/** Build absolute Edit Properties URL (PageType=6) */
function buildSingleEditUrl(siteUrl: string, listId: string, itemId: number, sourceAbs: string) {
	const src = `&Source=${encodeURIComponent(sourceAbs)}`;
	return `${siteUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${encodeURIComponent(
		listId
	)}&ID=${encodeURIComponent(itemId)}${src}`;
}

/** Bulk view URL with helper params we can act on via DOM nudges */
function buildBulkViewUrl(
	siteUrl: string,
	libraryServerRelativeUrl: string,
	ids: number[],
	viewId?: string
) {
	const idsCsv = ids.join(',');
	const view = viewId ? `?viewid=%7B${encodeURIComponent(viewId)}%7D` : '';
	const sep = view ? '&' : '?';
	return `${siteUrl}${libraryServerRelativeUrl}${view}${sep}selected=${encodeURIComponent(
		idsCsv
	)}&openPane=1`;
}

/** Best-effort: detect when single edit returns to Source (saved or canceled) */
function detectSingleSaveFromUrl(url: string, hostSource: string): boolean {
	try {
		if (!url.includes('Source=')) return false;
		return decodeURIComponent(url).includes(hostSource);
	} catch {
		return false;
	}
}

export const LibraryItemEditorLauncher: React.FC<LibraryItemEditorLauncherProps> = (props) => {
	const {
		siteUrl,
		spfxContext,
		libraryServerRelativeUrl,
		itemIds,
		viewId,

		renderMode,
		isOpen = true,

		onDetermined,
		onOpen,
		onSaved,
		onDismiss,

		enableBulkAutoRefresh = true,
		bulkWatchAllItems = true,

		disableDomNudges = false,
		sandboxExtra,

		autoHeightBestEffort = true,
	} = props;

	const [targetUrl, setTargetUrl] = useState<string>('');
	const [mode, setMode] = useState<'single' | 'bulk'>(() =>
		itemIds?.length === 1 ? 'single' : 'bulk'
	);
	const [modalOpen, setModalOpen] = useState<boolean>(isOpen && renderMode === 'modal');

	const sp = useMemo(() => spfi(siteUrl).using(PnP_SPFX(spfxContext)), [siteUrl, spfxContext]);
	const iframeRef = useRef<HTMLIFrameElement | null>(null);

	// keep modal state synced with prop
	useEffect(() => {
		if (renderMode === 'modal') setModalOpen(!!isOpen);
	}, [renderMode, isOpen]);

	// -------- determine the target URL (and fire onDetermined) --------
	useEffect(() => {
		let disposed = false;
		(async () => {
			if (!itemIds?.length) return;
			const single = itemIds.length === 1;
			setMode(single ? 'single' : 'bulk');

			try {
				if (single) {
					const listId = await resolveListId(sp, libraryServerRelativeUrl);
					const url = buildSingleEditUrl(siteUrl, listId, itemIds[0], window.location.href);
					if (disposed) return;

					setTargetUrl(url);
					onDetermined?.({ mode: 'single', url, bulk: false });

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener');
						onOpen?.({ mode: 'single', url });
						onDismiss?.();
					} else if (renderMode === 'samepage') {
						window.location.href = url;
						onOpen?.({ mode: 'single', url });
						onDismiss?.();
					}
				} else {
					const url = buildBulkViewUrl(siteUrl, libraryServerRelativeUrl, itemIds, viewId);
					if (disposed) return;

					setTargetUrl(url);
					onDetermined?.({ mode: 'bulk', url, bulk: true });

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener');
						onOpen?.({ mode: 'bulk', url });
						onDismiss?.();
					} else if (renderMode === 'samepage') {
						window.location.href = url;
						onOpen?.({ mode: 'bulk', url });
						onDismiss?.();
					}
				}
			} catch (e) {
				// eslint-disable-next-line no-console
				console.error('[LibraryItemEditorLauncher] determine URL failed', e);
				onDismiss?.();
			}
		})();
		return () => {
			disposed = true;
		};
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [siteUrl, libraryServerRelativeUrl, itemIds, viewId, renderMode]);

	// -------- bulk auto-close by Modified polling --------
	useEffect(() => {
		if (mode !== 'bulk' || !enableBulkAutoRefresh || renderMode !== 'modal') return;
		let timer: number | undefined;

		const idsToWatch = bulkWatchAllItems ? itemIds : [itemIds[0]];
		const original: Record<number, string> = {};
		let seeded = false;

		const tick = async () => {
			try {
				if (!seeded) {
					for (const id of idsToWatch) {
						const it: any = await sp.web
							.getList(libraryServerRelativeUrl)
							.items.getById(id)
							.select('Modified')();
						original[id] = it?.Modified as string;
					}
					seeded = true;
					return;
				}
				for (const id of idsToWatch) {
					const it: any = await sp.web
						.getList(libraryServerRelativeUrl)
						.items.getById(id)
						.select('Modified')();
					if (it?.Modified && original[id] && it.Modified !== original[id]) {
						setModalOpen(false);
						onSaved?.();
						onDismiss?.();
						if (timer) window.clearInterval(timer);
						return;
					}
				}
			} catch {
				/* ignore transient errors */
			}
		};

		timer = window.setInterval(tick, 5000);
		tick();

		return () => {
			if (timer) window.clearInterval(timer);
		};
	}, [
		mode,
		enableBulkAutoRefresh,
		bulkWatchAllItems,
		itemIds,
		renderMode,
		sp,
		libraryServerRelativeUrl,
		onSaved,
		onDismiss,
	]);

	// -------- iframe load: fire onOpen, detect saved, DOM nudges --------
	const onIframeLoad = () => {
		if (!iframeRef.current) return;
		const frame = iframeRef.current;

		// 🔵 onOpen fires when the frame loads a target
		if (targetUrl) onOpen?.({ mode, url: targetUrl });

		// Single-edit: detect return-to-Source as "saved/closed"
		try {
			const href = (frame.contentWindow as any)?.location?.href as string | undefined;
			if (href && mode === 'single' && detectSingleSaveFromUrl(href, window.location.href)) {
				setModalOpen(false);
				onSaved?.();
				onDismiss?.();
				return;
			}
		} catch {
			/* same-origin guarded; sandbox allows same-origin */
		}

		// Bulk DOM nudges: select rows & open details pane
		if (!disableDomNudges && mode === 'bulk') {
			try {
				const doc = frame.contentDocument || frame.contentWindow?.document;
				if (!doc) return;

				const openPane = () => {
					const btn = doc.querySelector(
						'[data-automationid="DetailsPaneButton"],button[name="OpenDetailsPane"]'
					) as HTMLButtonElement | null;
					btn?.click();
				};

				const selectByIds = (ids: number[]) => {
					const rows = Array.from(doc.querySelectorAll('[role="row"]')) as HTMLElement[];
					ids.forEach((id) => {
						const match = rows.find((r) => r.innerText?.match(new RegExp(`(^|\\s)${id}(\\s|$)`)));
						if (match) {
							const cb = match.querySelector('[role="checkbox"]') as HTMLElement | null;
							if (cb && cb.getAttribute('aria-checked') !== 'true') cb.click();
							else (match as any).click?.();
						}
					});
				};

				setTimeout(() => {
					selectByIds(itemIds);
					openPane();
				}, 400);
			} catch {
				/* ignore if DOM inaccessible */
			}
		}
	};

	// -------- render (modal only) --------
	if (renderMode === 'newtab' || renderMode === 'samepage') return null;

	const iframeStyle: React.CSSProperties = {
		width: '100%',
		border: 'none',
		height: autoHeightBestEffort
			? `${Math.max(500, Math.floor(window.innerHeight * 0.8))}px`
			: '80vh',
		overflow: 'hidden',
	};

	const sandbox = `allow-scripts allow-same-origin allow-forms allow-popups${
		sandboxExtra ? ` ${sandboxExtra}` : ''
	}`;

	return (
		<Dialog
			hidden={!modalOpen}
			onDismiss={() => {
				setModalOpen(false);
				onDismiss?.();
			}}
			dialogContentProps={{ type: DialogType.close, title: undefined }}
			minWidth="60%"
			maxWidth="98%"
			modalProps={{ isBlocking: true }}
		>
			<Stack horizontal horizontalAlign="end">
				<IconButton
					aria-label="Close"
					iconProps={{ iconName: 'Cancel' }}
					onClick={() => {
						setModalOpen(false);
						onDismiss?.();
					}}
				/>
			</Stack>

			{!!targetUrl && (
				<iframe
					ref={iframeRef}
					title="Edit properties"
					src={targetUrl} // absolute; not the current page
					style={iframeStyle}
					onLoad={onIframeLoad}
					sandbox={sandbox}
				/>
			)}
		</Dialog>
	);
};
