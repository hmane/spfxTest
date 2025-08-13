// src/webparts/UploadAndEdit/components/editor/LibraryItemEditorLauncher.tsx
import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog, DialogType } from '@fluentui/react';

import { spfi, SPFI } from '@pnp/sp';
import { SPFx as PnP_SPFX } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export type RenderMode = 'modal' | 'samepage' | 'newtab';

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
	siteUrl: string;
	spfxContext: any;

	libraryServerRelativeUrl: string;
	itemIds: number[];
	contentTypeId?: string;

	viewId?: string; // optional minimal view for bulk

	renderMode: RenderMode;
	isOpen?: boolean; // modal only

	onDetermined?: (info: LauncherDeterminedInfo) => void;
	onOpen?: (info: LauncherOpenInfo) => void;
	onSaved?: () => void;
	onDismiss?: () => void;

	enableBulkAutoRefresh?: boolean;
	bulkWatchAllItems?: boolean;

	disableDomNudges?: boolean;
	sandboxExtra?: string;

	// UI niceties
	autoHeightBestEffort?: boolean;

	// Optional chrome hiders
	hideBreadcrumbs?: boolean;
	hideContentTypeField?: boolean;
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

function isListFormEditUrl(href: string): boolean {
	try {
		const u = new URL(href, window.location.origin);
		return (
			u.pathname.toLowerCase().includes('/_layouts/15/listform.aspx') &&
			(u.search.toLowerCase().includes('pagetype=6') || u.searchParams.get('PageType') === '6')
		);
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

		hideBreadcrumbs = false,
		hideContentTypeField = false,
	} = props;

	const [targetUrl, setTargetUrl] = useState<string>('');
	const [mode, setMode] = useState<'single' | 'bulk'>(() =>
		itemIds?.length === 1 ? 'single' : 'bulk'
	);
	const [modalOpen, setModalOpen] = useState<boolean>(isOpen && renderMode === 'modal');

	const sp = useMemo(() => spfi(siteUrl).using(PnP_SPFX(spfxContext)), [siteUrl, spfxContext]);
	const iframeRef = useRef<HTMLIFrameElement | null>(null);

	// single-form load guard
	const singleInitialLoadSeenRef = useRef(false);

	useEffect(() => {
		if (renderMode === 'modal') setModalOpen(!!isOpen);
	}, [renderMode, isOpen]);

	// Reset guard when inputs change
	useEffect(() => {
		singleInitialLoadSeenRef.current = false;
	}, [itemIds, libraryServerRelativeUrl, siteUrl]);

	// Determine the target URL (and fire onDetermined)
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

	// Bulk auto-close by Modified polling
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

	// Iframe load: onOpen + single save detection + bulk DOM nudges + CSS injection
	const onIframeLoad = () => {
		const frame = iframeRef.current;
		if (!frame) return;

		if (targetUrl) onOpen?.({ mode, url: targetUrl });

		// SINGLE: detect post-save navigation
		if (mode === 'single') {
			try {
				const href = (frame.contentWindow as any)?.location?.href as string | undefined;
				if (href) {
					// First load is the edit form itself—don't close
					if (!singleInitialLoadSeenRef.current) {
						singleInitialLoadSeenRef.current = true;
					} else {
						// If we left the edit form and navigated to our Source (host page), close
						const leftEditForm = !isListFormEditUrl(href);
						const returnedToHost = decodeURIComponent(href).includes(window.location.href);
						if (leftEditForm && returnedToHost) {
							// Prevent showing the host page in the iframe while we close:
							try {
								frame.style.visibility = 'hidden';
								frame.src = 'about:blank';
							} catch {}
							setModalOpen(false);
							onSaved?.();
							onDismiss?.();
							return;
						}
					}
				}
			} catch {
				/* ignore */
			}
		}

		// Inject minimal CSS to hide chrome/breadcrumb/CT field (best effort)
		try {
			const doc = frame.contentDocument || frame.contentWindow?.document;
			if (doc) {
				const styleId = 'launcher-hide-chrome-css';
				if (!doc.getElementById(styleId)) {
					const css: string[] = [];
					if (props.hideBreadcrumbs) {
						css.push(
							'.od-TopNav, .od-TopBar, .ms-CommandBar, .ms-Breadcrumb { display:none !important; }',
							'[data-automationid="Breadcrumb"] { display:none !important; }'
						);
					}
					if (props.hideContentTypeField) {
						css.push(
							'div[data-field="ContentType"], div[aria-label="Content type"], label:contains("Content type") { display:none !important; }',
							'[data-automationid="ContentTypeSelector"] { display:none !important; }'
						);
					}
					if (css.length) {
						const style = doc.createElement('style');
						style.id = styleId;
						style.type = 'text/css';
						style.appendChild(doc.createTextNode(css.join('\n')));
						doc.head.appendChild(style);
					}
				}
			}
		} catch {
			/* ignore */
		}

		// BULK DOM nudges
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
				/* ignore */
			}
		}
	};

	// render (modal only)
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
			{!!targetUrl && (
				<iframe
					ref={iframeRef}
					title="Edit properties"
					src={targetUrl}
					style={iframeStyle}
					onLoad={onIframeLoad}
					sandbox={sandbox}
				/>
			)}
		</Dialog>
	);
};
