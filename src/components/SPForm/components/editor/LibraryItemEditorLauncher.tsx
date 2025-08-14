// src/webparts/UploadAndEdit/components/editor/LibraryItemEditorLauncher.tsx
// CLEAN FINAL VERSION (with loading/timeout fixes)

import * as React from 'react';
import { useEffect, useMemo, useRef, useState, useCallback } from 'react';
import {
	Dialog,
	DialogType,
	MessageBar,
	MessageBarType,
	Spinner,
	SpinnerSize,
	Stack,
	Text,
	IconButton,
	ProgressIndicator,
	IDialogContentProps,
} from '@fluentui/react';

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

	viewId?: string;
	renderMode: RenderMode;
	isOpen?: boolean;

	onDetermined?: (info: LauncherDeterminedInfo) => void;
	onOpen?: (info: LauncherOpenInfo) => void;
	onSaved?: () => void;
	onDismiss?: () => void;

	enableBulkAutoRefresh?: boolean;
	bulkWatchAllItems?: boolean;
	disableDomNudges?: boolean;
	sandboxExtra?: string;
	autoHeightBestEffort?: boolean;
	hideBreadcrumbs?: boolean;
	hideContentTypeField?: boolean;
}

interface ItemInfo {
	id: number;
	uniqueId: string;
	fileName: string;
	modified: string;
	serverRelativeUrl?: string;
	listItemId?: number;
}

interface LoadingState {
	isLoading: boolean;
	message: string;
	progress?: number;
}

/* ----------------------------- Utilities ----------------------------- */

const createAdvancedCSS = (hideBreadcrumbs: boolean, hideContentTypeField: boolean): string => {
	const css: string[] = [];

	css.push(`
    body { margin: 0 !important; padding: 0 !important; overflow-x: hidden !important; }
    .ms-Dialog-main { padding: 0 !important; }
    html, body { scroll-behavior: smooth; }
    .ms-Overlay--dark { background-color: rgba(0,0,0,0.1) !important; }
  `);

	if (hideBreadcrumbs) {
		css.push(`
      .od-TopNav, .od-TopBar, .ms-CommandBar, .ms-Breadcrumb, .ms-Nav,
      [data-automationid="Breadcrumb"], [data-automationid="breadcrumb"],
      [class*="breadcrumb" i], [class*="topnav" i], [class*="navigation" i],
      .od-SuiteNav, .suite-nav, .ms-NavBar, .od-AppBreadcrumb,
      nav[role="navigation"], .sp-appBar, .spPageChromeAppDiv,
      #spPageChromeAppDiv, .od-Shell-topBar, .od-TopBar-container,
      .ms-FocusZone[data-focuszone-id*="CommandBar"],
      [class*="CommandBar"][class*="breadcrumb" i],
      div[class*="topBar" i], div[class*="header" i]:not(.ms-Panel-header):not(.ms-Dialog-header),
      .ms-CommandBar--fixed, div[data-sp-feature-tag="Site Navigation"],
      div[data-sp-feature-tag="Top Navigation"], .spSiteHeader, .sp-siteHeader,
      #SuiteNavWrapper, #suiteBarDelta, .ms-srch-sb, .ms-core-pageTitle,
      [data-automation-id="contentHeader"], .sp-contentHeader,
      #s4-titlerow, #s4-ribbonrow, .ms-dlgFrame .ms-dlgTitleText
      { display:none !important; height:0 !important; overflow:hidden !important; visibility:hidden !important; }

      .od-Canvas, .Canvas, main[role="main"], .ms-Fabric, .root-40
      { margin-top:0 !important; padding-top:8px !important; }
    `);
	}

	if (hideContentTypeField) {
		css.push(`
      div[data-field="ContentType"], div[data-field="contenttype"],
      div[aria-label*="Content type" i], div[aria-label*="ContentType" i],
      [data-automationid="ContentTypeSelector"], [data-automationid="contenttypeselector"],
      .ms-TextField[aria-label*="Content Type" i], .ms-TextField[aria-label*="ContentType" i],
      input[aria-label*="Content Type" i], input[aria-label*="ContentType" i],
      .ms-FieldLabel[for*="ContentType" i], .ms-FieldLabel[for*="content-type" i],
      label[for*="ContentType" i], label[for*="content-type" i],
      .ms-Dropdown[aria-label*="Content Type" i], .ms-ComboBox[aria-label*="Content Type" i],
      tr:has(td[data-field="ContentType"]), tr:has(.ms-FieldLabel[for*="ContentType" i]),
      .ms-FormField:has([data-field="ContentType"]), .ms-FormField:has([aria-label*="Content Type" i]),
      div[class*="field" i]:has([aria-label*="Content Type" i]),
      div[class*="control" i]:has([aria-label*="Content Type" i]),
      .propertyPane [aria-label*="Content Type" i], .ms-Panel [aria-label*="Content Type" i],
      [data-testid*="ContentType" i], [data-testid*="content-type" i]
      { display:none !important; height:0 !important; overflow:hidden !important; visibility:hidden !important; }
    `);
	}

	css.push(`
    .ms-Spinner { margin: 8px auto !important; }
    .ms-Panel-main, .ms-Dialog-main { padding: 12px !important; }
    .ms-Panel-footer .ms-Button, .ms-Dialog-actionsRight .ms-Button { margin: 0 4px !important; }
    .ms-DetailsList { margin-top: 8px !important; }
    .ms-MessageBar { margin: 8px 0 !important; border-radius: 4px !important; }
    .ms-Fabric :focus { outline: 2px solid #0078d4 !important; outline-offset: 2px !important; }
    *, *::before, *::after { animation-duration: 0.01ms !important; animation-iteration-count: 1 !important; transition-duration: 0.01ms !important; }
    @media (max-width:768px){ .ms-Panel-main, .ms-Dialog-main { padding: 8px !important; } }
  `);

	return css.join('\n');
};

async function resolveListId(sp: SPFI, libraryServerRelativeUrl: string): Promise<string> {
	let lastError: Error | null = null;
	for (let attempt = 0; attempt < 3; attempt++) {
		try {
			const list = await sp.web.getList(libraryServerRelativeUrl).select('Id')();
			const listId = list?.Id as string;
			if (!listId) throw new Error('List ID not found in response');
			return listId;
		} catch (err) {
			lastError = err as Error;
			if (attempt < 2) await new Promise((r) => setTimeout(r, 1000 * (attempt + 1)));
		}
	}
	throw new Error(
		`Unable to access library: ${lastError instanceof Error ? lastError.message : 'Unknown error'}`
	);
}

async function getItemDetails(
	sp: SPFI,
	libraryServerRelativeUrl: string,
	itemIds: number[]
): Promise<ItemInfo[]> {
	const list = sp.web.getList(libraryServerRelativeUrl);
	const results = await Promise.allSettled(
		itemIds.map(async (id) => {
			const item: any = await list.items
				.getById(id)
				.select('Id', 'UniqueId', 'FileLeafRef', 'Modified', 'FileRef', 'GUID')();
			return {
				id: item.Id,
				uniqueId: item.UniqueId || item.GUID,
				fileName: item.FileLeafRef,
				modified: item.Modified,
				serverRelativeUrl: item.FileRef,
				listItemId: item.Id,
			};
		})
	);
	const ok: ItemInfo[] = [];
	for (const r of results) if (r.status === 'fulfilled') ok.push(r.value);
	return ok;
}

function buildSingleEditUrl(
	siteUrl: string,
	listId: string,
	itemId: number,
	sourceAbs: string
): string {
	const cleanSiteUrl = siteUrl.replace(/\/$/, '');
	const src = `&Source=${encodeURIComponent(sourceAbs)}`;
	return `${cleanSiteUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${encodeURIComponent(
		listId
	)}&ID=${encodeURIComponent(itemId)}${src}`;
}

function buildBulkViewUrl(
	siteUrl: string,
	libraryServerRelativeUrl: string,
	itemDetails: ItemInfo[],
	viewId?: string
): string {
	const cleanSiteUrl = siteUrl.replace(/\/$/, '');
	let url = `${cleanSiteUrl}${libraryServerRelativeUrl}`;
	const params = new URLSearchParams();

	if (viewId) {
		const cleanViewId = viewId.replace(/[{}]/g, '');
		params.set('viewid', `{${cleanViewId}}`);
	}

	if (itemDetails.length > 0) {
		const itemIds = itemDetails.map((i) => i.id.toString());
		const fileNames = itemDetails.map((i) => i.fileName).join(',');

		params.set('FilterField1', 'FileLeafRef');
		params.set('FilterValue1', fileNames);
		params.set('FilterType1', 'Text');
		params.set('env', 'WebView');
		params.set('OR', 'Teams-HL');
		params.set('selectedItems', itemIds.join(','));
	}

	const q = params.toString();
	if (q) url += `?${q}`;
	return url;
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

/* ----------------------------- Component ----------------------------- */

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
	const [itemDetails, setItemDetails] = useState<ItemInfo[]>([]);
	const [loadingState, setLoadingState] = useState<LoadingState>({ isLoading: false, message: '' });
	const [error, setError] = useState<string | null>(null);
	const [cssInjected, setCssInjected] = useState<boolean>(false);
	const [isInitialized, setIsInitialized] = useState<boolean>(false);

	const sp = useMemo(() => spfi(siteUrl).using(PnP_SPFX(spfxContext)), [siteUrl, spfxContext]);
	const iframeRef = useRef<HTMLIFrameElement | null>(null);
	const singleInitialLoadSeenRef = useRef(false);
	const changeDetectionTimer = useRef<number>();
	const loadingTimeoutRef = useRef<number>();
	const initializationRef = useRef<boolean>(false);

	const stableItemIds = useMemo(() => itemIds.join(','), [itemIds]);
	const stableKey = useMemo(
		() => `${siteUrl}|${libraryServerRelativeUrl}|${stableItemIds}|${viewId || ''}|${renderMode}`,
		[siteUrl, libraryServerRelativeUrl, stableItemIds, viewId, renderMode]
	);

	const injectAdvancedCSS = useCallback(
		(doc: Document) => {
			if (cssInjected) return;
			try {
				const styleId = 'enhanced-launcher-styles';
				if (!doc.getElementById(styleId)) {
					const cssContent = createAdvancedCSS(hideBreadcrumbs, hideContentTypeField);
					const style = doc.createElement('style');
					style.id = styleId;
					style.type = 'text/css';
					style.appendChild(doc.createTextNode(cssContent));
					(doc.head || doc.documentElement).appendChild(style);
					setCssInjected(true);
				}
			} catch (e) {
				console.warn('CSS injection failed:', e);
			}
		},
		[hideBreadcrumbs, hideContentTypeField, cssInjected]
	);

	const performAdvancedBulkSelection = useCallback(
		(doc: Document, attempt = 0) => {
			const maxAttempts = 20;
			if (attempt >= maxAttempts) {
				setLoadingState({ isLoading: false, message: '' });
				return;
			}

			try {
				setLoadingState({
					isLoading: true,
					message: 'Selecting files…',
					progress: (attempt / maxAttempts) * 100,
				});

				if (doc.readyState !== 'complete' || !doc.body) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 800);
					return;
				}

				const selectors = {
					listItems: [
						'[data-automationid="DetailsRow"]',
						'[role="row"][data-selection-index]',
						'[data-list-index]',
						'.ms-DetailsRow',
					],
					checkboxes: [
						'[data-selection-toggle="true"]',
						'button[data-selection-toggle="true"]',
						'[role="checkbox"]',
						'[data-automationid="DetailsRowCheck"]',
						'.ms-Check input[type="checkbox"]',
						'[aria-label*="Select row"]',
						'[data-automationid="SelectionCheckbox"]',
					],
					detailsPane: [
						'[data-automationid="PropertyPaneButton"]',
						'[data-automationid="InfoButton"]',
						'[aria-label*="Information panel"]',
						'[aria-label*="Details panel"]',
						'[aria-label*="Properties"]',
					],
				};

				let selectedCount = 0;

				for (const sel of selectors.listItems) {
					const rows = Array.from(doc.querySelectorAll(sel));
					if (!rows.length) continue;

					itemDetails.forEach((item) => {
						const fn = item.fileName;
						const base = fn.replace(/\.[^/.]+$/, '');
						let itemSelected = false;

						for (const row of rows as HTMLElement[]) {
							if (itemSelected) break;
							const text = row.innerText || row.textContent || '';
							const match =
								text.includes(fn) ||
								text.includes(base) ||
								row.querySelector(`[title*="${fn}"]`) ||
								row.querySelector(`[aria-label*="${fn}"]`) ||
								row.getAttribute('data-item-id') === String(item.id) ||
								row.getAttribute('data-unique-id') === item.uniqueId;

							if (match) {
								for (const cbSel of selectors.checkboxes) {
									const cb = row.querySelector(cbSel) as HTMLElement | null;
									if (!cb) continue;
									const already =
										cb.getAttribute('aria-checked') === 'true' ||
										cb.getAttribute('checked') === 'true' ||
										row.getAttribute('aria-selected') === 'true' ||
										row.classList.contains('is-selected') ||
										row.classList.contains('ms-DetailsRow--selected');

									if (!already) {
										try {
											cb.click();
											['mousedown', 'mouseup', 'click'].forEach((type) =>
												cb.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true }))
											);
										} catch {}
									}
									selectedCount++;
									itemSelected = true;
									break;
								}

								if (!itemSelected) {
									try {
										(row as HTMLElement).click();
										selectedCount++;
										itemSelected = true;
									} catch {}
								}
							}
						}
					});

					if (selectedCount > 0) break;
				}

				// Open details pane if something was selected
				if (selectedCount > 0) {
					setTimeout(() => {
						for (const s of selectors.detailsPane) {
							const btn = doc.querySelector(s) as HTMLElement | null;
							if (btn) {
								try {
									btn.click();
									break;
								} catch {}
							}
						}
					}, 1200);
				}

				if (selectedCount < itemDetails.length && attempt < maxAttempts - 1) {
					const delay = Math.min(3000, 1000 + attempt * 300);
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), delay);
				} else {
					// selection finished (full or partial) — remove overlay
					setLoadingState({ isLoading: false, message: '' });
				}
			} catch (e) {
				if (attempt < maxAttempts - 1) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 2000);
				} else {
					setLoadingState({ isLoading: false, message: '' });
				}
			}
		},
		[itemDetails]
	);

	const onIframeLoad = useCallback(() => {
		const frame = iframeRef.current;
		if (!frame) return;

		// Any pending init timeout should be cleared once something loads
		if (loadingTimeoutRef.current) {
			clearTimeout(loadingTimeoutRef.current);
			loadingTimeoutRef.current = undefined as any;
		}

		if (mode === 'single') {
			// hide any loader for single as soon as the form is available
			setLoadingState({ isLoading: false, message: '' });
		}

		if (targetUrl && onOpen) onOpen({ mode, url: targetUrl });

		const doc = frame.contentDocument || frame.contentWindow?.document;
		if (doc) {
			const inject = () => {
				try {
					injectAdvancedCSS(doc);
				} catch {}
			};
			if (doc.readyState === 'complete') inject();
			else {
				doc.addEventListener('DOMContentLoaded', inject);
				setTimeout(inject, 1000);
			}
		}

		// Detect single edit completion by navigation back to Source
		if (mode === 'single') {
			try {
				const href = (frame.contentWindow as any)?.location?.href as string | undefined;
				if (href) {
					if (!singleInitialLoadSeenRef.current) {
						singleInitialLoadSeenRef.current = true;
					} else {
						const leftEditForm = !isListFormEditUrl(href);
						const returnedToHost = decodeURIComponent(href).includes(window.location.href);
						if (leftEditForm && returnedToHost) {
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
			} catch {}
		}

		// Bulk: show overlay only while auto-selecting (if enabled)
		if (!disableDomNudges && mode === 'bulk' && itemDetails.length > 0 && doc) {
			setLoadingState({ isLoading: true, message: 'Selecting files…', progress: 0 });
			setTimeout(() => performAdvancedBulkSelection(doc, 0), 2000);
		} else if (mode === 'bulk') {
			setLoadingState({ isLoading: false, message: '' });
		}
	}, [
		mode,
		targetUrl,
		itemDetails.length,
		disableDomNudges,
		injectAdvancedCSS,
		performAdvancedBulkSelection,
		onOpen,
		onSaved,
		onDismiss,
	]);

	/* ------------------------------- Effects ------------------------------ */

	// Cleanup timers on unmount
	useEffect(() => {
		return () => {
			if (changeDetectionTimer.current) clearInterval(changeDetectionTimer.current);
			if (loadingTimeoutRef.current) clearTimeout(loadingTimeoutRef.current);
		};
	}, []);

	// Modal visibility driven by prop
	useEffect(() => {
		if (renderMode === 'modal') setModalOpen(!!isOpen);
	}, [renderMode, isOpen]);

	// MAIN initialization (build URL; no long-lived loader after URL is ready)
	useEffect(() => {
		if (!itemIds?.length || isInitialized) return;
		if (initializationRef.current) return;

		initializationRef.current = true;
		let disposed = false;

		const initializeEditor = async () => {
			try {
				setError(null);
				setMode(itemIds.length === 1 ? 'single' : 'bulk');
				setLoadingState({ isLoading: true, message: 'Preparing editor…', progress: 0 });

				if (itemIds.length === 1) {
					// SINGLE
					setLoadingState({
						isLoading: true,
						message: 'Resolving list information…',
						progress: 20,
					});
					const listId = await Promise.race([
						resolveListId(sp, libraryServerRelativeUrl),
						new Promise<never>((_, reject) =>
							setTimeout(() => reject(new Error('List ID resolution timeout')), 15000)
						),
					]);
					if (disposed) return;

					setLoadingState({ isLoading: true, message: 'Building edit form URL…', progress: 60 });
					const url = buildSingleEditUrl(siteUrl, listId, itemIds[0], window.location.href);
					if (disposed) return;

					setTargetUrl(url);
					// URL ready → loader off; iframe onLoad will manage any further UI
					setLoadingState({ isLoading: false, message: '' });
					onDetermined?.({ mode: 'single', url, bulk: false });

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener,noreferrer');
						onOpen?.({ mode: 'single', url });
						onDismiss?.();
						return;
					} else if (renderMode === 'samepage') {
						onOpen?.({ mode: 'single', url });
						window.location.href = url;
						onDismiss?.();
						return;
					}
				} else {
					// BULK
					setLoadingState({ isLoading: true, message: 'Retrieving file details…', progress: 20 });
					const details = await Promise.race([
						getItemDetails(sp, libraryServerRelativeUrl, itemIds),
						new Promise<never>((_, reject) =>
							setTimeout(() => reject(new Error('Item details timeout')), 20000)
						),
					]);
					if (disposed) return;

					if (!details.length) throw new Error('Could not retrieve details for selected files.');

					setItemDetails(details);
					setLoadingState({ isLoading: true, message: 'Building bulk edit URL…', progress: 60 });
					const url = buildBulkViewUrl(siteUrl, libraryServerRelativeUrl, details, viewId);
					if (disposed) return;

					setTargetUrl(url);
					setLoadingState({ isLoading: false, message: '' });
					onDetermined?.({ mode: 'bulk', url, bulk: true });

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener,noreferrer');
						onOpen?.({ mode: 'bulk', url });
						onDismiss?.();
						return;
					} else if (renderMode === 'samepage') {
						onOpen?.({ mode: 'bulk', url });
						window.location.href = url;
						onDismiss?.();
						return;
					}
				}

				setIsInitialized(true);
			} catch (e) {
				if (!disposed) {
					const msg = e instanceof Error ? e.message : 'Failed to initialize editor';
					setError(msg);
					setLoadingState({ isLoading: false, message: '' });
				}
			} finally {
				if (!disposed) initializationRef.current = false;
			}
		};

		// Safety timeout during init — will be cleared on iframe load or once URL is set
		loadingTimeoutRef.current = window.setTimeout(() => {
			setLoadingState({ isLoading: false, message: '' });
			setError('Loading took too long. Please try again.');
		}, 30000);

		const initTimer = setTimeout(() => {
			if (!disposed) void initializeEditor();
		}, 100);

		return () => {
			disposed = true;
			clearTimeout(initTimer);
			if (loadingTimeoutRef.current) {
				clearTimeout(loadingTimeoutRef.current);
				loadingTimeoutRef.current = undefined as any;
			}
		};
	}, [
		itemIds,
		libraryServerRelativeUrl,
		siteUrl,
		viewId,
		renderMode,
		isInitialized,
		sp,
		onDetermined,
		onOpen,
		onDismiss,
	]);

	// Reset when key props change
	useEffect(() => {
		setIsInitialized(false);
		initializationRef.current = false;
		singleInitialLoadSeenRef.current = false;
		setCssInjected(false);
		setError(null);
		setTargetUrl('');
		setLoadingState({ isLoading: false, message: '' });

		if (changeDetectionTimer.current) clearInterval(changeDetectionTimer.current);
		if (loadingTimeoutRef.current) {
			clearTimeout(loadingTimeoutRef.current);
			loadingTimeoutRef.current = undefined as any;
		}
	}, [stableKey]);

	// Auto-close on item change (bulk in modal)
	useEffect(() => {
		if (mode !== 'bulk' || !enableBulkAutoRefresh || renderMode !== 'modal' || !isInitialized)
			return;

		const idsToWatch = bulkWatchAllItems ? itemIds : [itemIds[0]];
		const originalModified: Record<number, string> = {};
		let initialized = false;

		const checkForChanges = async () => {
			try {
				if (!initialized) {
					for (const id of idsToWatch) {
						const item: any = await sp.web
							.getList(libraryServerRelativeUrl)
							.items.getById(id)
							.select('Modified')();
						originalModified[id] = item?.Modified as string;
					}
					initialized = true;
					return;
				}
				for (const id of idsToWatch) {
					const item: any = await sp.web
						.getList(libraryServerRelativeUrl)
						.items.getById(id)
						.select('Modified')();
					if (item?.Modified && originalModified[id] && item.Modified !== originalModified[id]) {
						setModalOpen(false);
						onSaved?.();
						onDismiss?.();
						return;
					}
				}
			} catch {}
		};

		changeDetectionTimer.current = window.setInterval(checkForChanges, 3000);
		void checkForChanges();

		return () => {
			if (changeDetectionTimer.current) clearInterval(changeDetectionTimer.current);
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
		isInitialized,
	]);

	/* -------------------------------- Render ------------------------------- */

	if (renderMode === 'newtab' || renderMode === 'samepage') return null;

	const dialogContentProps: IDialogContentProps = {
		type: DialogType.close,
		title:
			mode === 'bulk'
				? `Edit Properties - ${itemIds.length} file${itemIds.length > 1 ? 's' : ''}`
				: 'Edit File Properties',
		showCloseButton: true,
		styles: {
			content: { padding: '0', margin: '0' },
			header: { padding: '12px 20px 8px', borderBottom: '1px solid #edebe9' },
			title: { fontSize: '18px', fontWeight: 600 },
		},
	};

	const iframeStyle: React.CSSProperties = {
		width: '100%',
		border: 'none',
		height: autoHeightBestEffort
			? `${Math.max(600, Math.floor(window.innerHeight * 0.85))}px`
			: '85vh',
		overflow: 'hidden',
		// hide iframe only for single-mode while "preparing" (before onLoad fires)
		display: loadingState.isLoading && mode === 'single' ? 'none' : 'block',
	};

	const sandbox = `allow-scripts allow-same-origin allow-forms allow-popups allow-downloads allow-modals allow-presentation${
		sandboxExtra ? ` ${sandboxExtra}` : ''
	}`;

	return (
		<Dialog
			hidden={!modalOpen}
			onDismiss={() => {
				setModalOpen(false);
				onDismiss?.();
			}}
			dialogContentProps={dialogContentProps}
			minWidth="75%"
			maxWidth="98%"
			modalProps={{
				isBlocking: true,
				styles: {
					main: { maxHeight: '95vh', height: 'auto', padding: 0, margin: 0 },
					scrollableContent: { padding: 0, margin: 0, overflow: 'hidden' },
				},
			}}
		>
			<Stack tokens={{ childrenGap: 0 }} styles={{ root: { height: '100%' } }}>
				{/* Small pre-load indicator only before URL is ready (we keep this minimal) */}
				{loadingState.isLoading && mode === 'single' && !targetUrl && (
					<Stack
						tokens={{ childrenGap: 12 }}
						styles={{
							root: { padding: 20, minHeight: 160, justifyContent: 'center', alignItems: 'center' },
						}}
					>
						<Spinner size={SpinnerSize.large} />
						<Text variant="medium" styles={{ root: { textAlign: 'center' } }}>
							{loadingState.message || 'Preparing…'}
						</Text>
						{typeof loadingState.progress === 'number' && (
							<ProgressIndicator
								percentComplete={loadingState.progress / 100}
								description={`${Math.round(loadingState.progress)}%`}
								styles={{ root: { width: 300 } }}
							/>
						)}
					</Stack>
				)}

				{error && (
					<Stack styles={{ root: { padding: 16 } }}>
						<MessageBar
							messageBarType={MessageBarType.error}
							onDismiss={() => setError(null)}
							actions={
								<div>
									<IconButton
										iconProps={{ iconName: 'Refresh' }}
										title="Retry"
										ariaLabel="Retry"
										onClick={() => {
											setError(null);
											setLoadingState({ isLoading: false, message: '' });
											window.location.reload();
										}}
									/>
								</div>
							}
						>
							<strong>Error:</strong> {error}
						</MessageBar>
					</Stack>
				)}

				{/* Bulk header info (when visible) */}
				{mode === 'bulk' && !loadingState.isLoading && !error && itemDetails.length > 0 && (
					<Stack
						horizontal
						horizontalAlign="space-between"
						verticalAlign="center"
						styles={{
							root: {
								padding: '8px 16px',
								backgroundColor: '#f8f9fa',
								borderBottom: '1px solid #edebe9',
								fontSize: 13,
							},
						}}
					>
						<Text variant="small">
							<strong>{itemDetails.length}</strong> file{itemDetails.length > 1 ? 's' : ''} selected
							for bulk editing
						</Text>
						<Stack horizontal tokens={{ childrenGap: 8 }}>
							{enableBulkAutoRefresh && (
								<Text variant="small" styles={{ root: { color: '#0078d4' } }}>
									● Auto-save detection enabled
								</Text>
							)}
							<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
								{new Date().toLocaleTimeString()}
							</Text>
						</Stack>
					</Stack>
				)}

				{/* Main iframe */}
				{targetUrl && !error && (
					<div style={{ flex: 1, position: 'relative', overflow: 'hidden' }}>
						<iframe
							ref={iframeRef}
							title={mode === 'bulk' ? 'Bulk Edit Properties' : 'Edit File Properties'}
							src={targetUrl}
							style={iframeStyle}
							onLoad={onIframeLoad}
							onError={() => {
								setError('Failed to load the edit form. Please try again or refresh the page.');
								setLoadingState({ isLoading: false, message: '' });
							}}
							sandbox={sandbox}
							loading="lazy"
						/>

						{/* Bulk overlay ONLY while auto-selecting */}
						{loadingState.isLoading && mode === 'bulk' && (
							<div
								style={{
									position: 'absolute',
									inset: 0,
									backgroundColor: 'rgba(255,255,255,0.9)',
									display: 'flex',
									alignItems: 'center',
									justifyContent: 'center',
									zIndex: 1000,
								}}
							>
								<Stack tokens={{ childrenGap: 12 }} horizontalAlign="center">
									<Spinner size={SpinnerSize.large} />
									<Text variant="medium">{loadingState.message || 'Working…'}</Text>
									{typeof loadingState.progress === 'number' && (
										<ProgressIndicator
											percentComplete={loadingState.progress / 100}
											styles={{ root: { width: 200 } }}
										/>
									)}
								</Stack>
							</div>
						)}
					</div>
				)}

				{/* Footer */}
				<Stack
					horizontal
					horizontalAlign="space-between"
					verticalAlign="center"
					styles={{
						root: {
							padding: '8px 16px',
							borderTop: '1px solid #edebe9',
							backgroundColor: '#fafafa',
							fontSize: 12,
							color: '#605e5c',
						},
					}}
				>
					<Text variant="small">
						{mode === 'single'
							? 'Edit the properties and click Save to apply changes'
							: 'Select files and use the details pane to edit properties in bulk'}
					</Text>
					<Stack horizontal tokens={{ childrenGap: 12 }}>
						{process.env.NODE_ENV === 'development' && (
							<Text variant="small">Debug: {mode} mode</Text>
						)}
						<Text variant="small">Press Esc to close</Text>
					</Stack>
				</Stack>
			</Stack>
		</Dialog>
	);
};
