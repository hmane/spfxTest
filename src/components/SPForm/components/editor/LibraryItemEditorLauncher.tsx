// Enhanced LibraryItemEditorLauncher with better CSS targeting and UX
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
}

interface LoadingState {
	isLoading: boolean;
	message: string;
	progress?: number;
}

/** Enhanced CSS injection with comprehensive selectors */
const createAdvancedCSS = (hideBreadcrumbs: boolean, hideContentTypeField: boolean): string => {
	const css: string[] = [];

	// Base improvements for better UX
	css.push(`
		/* Remove unnecessary margins and padding for cleaner look */
		body {
			margin: 0 !important;
			padding: 0 !important;
			overflow-x: hidden !important;
		}

		/* Improve overall spacing */
		.ms-Dialog-main {
			padding: 0 !important;
		}

		/* Better scrolling */
		html, body {
			scroll-behavior: smooth;
		}

		/* Hide loading overlays that might interfere */
		.ms-Overlay--dark {
			background-color: rgba(0, 0, 0, 0.1) !important;
		}
	`);

	if (hideBreadcrumbs) {
		css.push(`
			/* Hide all possible breadcrumb variations */
			.od-TopNav,
			.od-TopBar,
			.ms-CommandBar,
			.ms-Breadcrumb,
			.ms-Nav,
			[data-automationid="Breadcrumb"],
			[data-automationid="breadcrumb"],
			[class*="breadcrumb" i],
			[class*="topnav" i],
			[class*="navigation" i],
			.od-SuiteNav,
			.suite-nav,
			.ms-NavBar,
			.od-AppBreadcrumb,
			nav[role="navigation"],
			.sp-appBar,
			.spPageChromeAppDiv,
			#spPageChromeAppDiv,
			.od-Shell-topBar,
			.od-TopBar-container,
			.ms-FocusZone[data-focuszone-id*="CommandBar"],
			[class*="CommandBar"][class*="breadcrumb" i],
			div[class*="topBar" i],
			div[class*="header" i]:not(.ms-Panel-header):not(.ms-Dialog-header),
			.ms-CommandBar--fixed,
			/* Modern SharePoint specific */
			div[data-sp-feature-tag="Site Navigation"],
			div[data-sp-feature-tag="Top Navigation"],
			.spSiteHeader,
			.sp-siteHeader,
			/* Remove suite bar */
			#SuiteNavWrapper,
			#suiteBarDelta,
			.ms-srch-sb,
			/* Remove page title area */
			.ms-core-pageTitle,
			[data-automation-id="contentHeader"],
			.sp-contentHeader,
			/* Hide the entire chrome area */
			#s4-titlerow,
			#s4-ribbonrow,
			.ms-dlgFrame .ms-dlgTitleText
			{
				display: none !important;
				height: 0 !important;
				overflow: hidden !important;
				visibility: hidden !important;
			}

			/* Adjust page layout after hiding navigation */
			.od-Canvas,
			.Canvas,
			main[role="main"],
			.ms-Fabric,
			.root-40 {
				margin-top: 0 !important;
				padding-top: 8px !important;
			}
		`);
	}

	if (hideContentTypeField) {
		css.push(`
			/* Hide content type field variations */
			div[data-field="ContentType"],
			div[data-field="contenttype"],
			div[aria-label*="Content type" i],
			div[aria-label*="ContentType" i],
			[data-automationid="ContentTypeSelector"],
			[data-automationid="contenttypeselector"],
			.ms-TextField[aria-label*="Content Type" i],
			.ms-TextField[aria-label*="ContentType" i],
			input[aria-label*="Content Type" i],
			input[aria-label*="ContentType" i],
			/* Form field containers */
			.ms-FieldLabel[for*="ContentType" i],
			.ms-FieldLabel[for*="content-type" i],
			label[for*="ContentType" i],
			label[for*="content-type" i],
			/* Dropdown variations */
			.ms-Dropdown[aria-label*="Content Type" i],
			.ms-ComboBox[aria-label*="Content Type" i],
			/* Field containers in forms */
			tr:has(td[data-field="ContentType"]),
			tr:has(.ms-FieldLabel[for*="ContentType" i]),
			.ms-FormField:has([data-field="ContentType"]),
			.ms-FormField:has([aria-label*="Content Type" i]),
			/* Modern form selectors */
			div[class*="field" i]:has([aria-label*="Content Type" i]),
			div[class*="control" i]:has([aria-label*="Content Type" i]),
			/* Property pane content type selectors */
			.propertyPane [aria-label*="Content Type" i],
			.ms-Panel [aria-label*="Content Type" i],
			/* Additional modern selectors */
			[data-testid*="ContentType" i],
			[data-testid*="content-type" i],
			div:has(> label:contains("Content Type")),
			div:has(> label:contains("ContentType"))
			{
				display: none !important;
				height: 0 !important;
				overflow: hidden !important;
				visibility: hidden !important;
			}
		`);
	}

	// Additional UX improvements
	css.push(`
		/* Better loading states */
		.ms-Spinner {
			margin: 8px auto !important;
		}

		/* Improve form layout */
		.ms-Panel-main,
		.ms-Dialog-main {
			padding: 12px !important;
		}

		/* Better button spacing */
		.ms-Panel-footer .ms-Button,
		.ms-Dialog-actionsRight .ms-Button {
			margin: 0 4px !important;
		}

		/* Improve table/list layouts */
		.ms-DetailsList {
			margin-top: 8px !important;
		}

		/* Better error/message bar styling */
		.ms-MessageBar {
			margin: 8px 0 !important;
			border-radius: 4px !important;
		}

		/* Improve focus indicators */
		.ms-Fabric :focus {
			outline: 2px solid #0078d4 !important;
			outline-offset: 2px !important;
		}

		/* Remove unnecessary animations that might cause issues */
		*, *::before, *::after {
			animation-duration: 0.01ms !important;
			animation-iteration-count: 1 !important;
			transition-duration: 0.01ms !important;
		}

		/* Ensure iframe content is responsive */
		@media (max-width: 768px) {
			.ms-Panel-main,
			.ms-Dialog-main {
				padding: 8px !important;
			}
		}
	`);

	return css.join('\n');
};

/** Resolve list GUID and get enhanced item details */
async function resolveListId(sp: SPFI, libraryServerRelativeUrl: string): Promise<string> {
	const list = await sp.web.getList(libraryServerRelativeUrl).select('Id')();
	return list?.Id as string;
}

/** Get comprehensive item details for better bulk operations */
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
				.select('Id', 'UniqueId', 'FileLeafRef', 'Modified', 'FileRef')();
			return {
				id: item.Id,
				uniqueId: item.UniqueId,
				fileName: item.FileLeafRef,
				modified: item.Modified,
				serverRelativeUrl: item.FileRef,
			};
		})
	);

	// Filter successful results and extract values
	const successfulResults: ItemInfo[] = [];
	for (const result of results) {
		if (result.status === 'fulfilled') {
			successfulResults.push(result.value);
		} else {
			console.warn('Failed to get item details:', result.reason);
		}
	}

	return successfulResults;
}

/** Build enhanced URLs for better SharePoint integration */
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
	const cleanLibraryUrl = libraryServerRelativeUrl.startsWith('/')
		? libraryServerRelativeUrl
		: `/${libraryServerRelativeUrl}`;

	let url = `${cleanSiteUrl}${cleanLibraryUrl}`;

	// Add view if specified
	if (viewId) {
		const cleanViewId = viewId.replace(/[{}]/g, '');
		url += `?viewid={${cleanViewId}}`;
	}

	// Add file selection parameters
	if (itemDetails.length > 0) {
		const separator = viewId ? '&' : '?';
		const fileParams = itemDetails
			.map((item) => `id=${encodeURIComponent(item.uniqueId)}`)
			.join('&');
		url += `${separator}${fileParams}&openPane=1`;
	}

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

	// Enhanced state management
	const [targetUrl, setTargetUrl] = useState<string>('');
	const [mode, setMode] = useState<'single' | 'bulk'>(() =>
		itemIds?.length === 1 ? 'single' : 'bulk'
	);
	const [modalOpen, setModalOpen] = useState<boolean>(isOpen && renderMode === 'modal');
	const [itemDetails, setItemDetails] = useState<ItemInfo[]>([]);
	const [loadingState, setLoadingState] = useState<LoadingState>({
		isLoading: true,
		message: 'Initializing...',
	});
	const [error, setError] = useState<string | null>(null);
	const [cssInjected, setCssInjected] = useState<boolean>(false);

	// Refs for managing lifecycle
	const sp = useMemo(() => spfi(siteUrl).using(PnP_SPFX(spfxContext)), [siteUrl, spfxContext]);
	const iframeRef = useRef<HTMLIFrameElement | null>(null);
	const singleInitialLoadSeenRef = useRef(false);
	const domNudgeAttempts = useRef(0);
	const changeDetectionTimer = useRef<number>();

	// Enhanced CSS injection with comprehensive selectors
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

					// Try multiple injection points
					if (doc.head) {
						doc.head.appendChild(style);
					} else if (doc.documentElement) {
						doc.documentElement.appendChild(style);
					}

					setCssInjected(true);
					console.log('Advanced CSS injected successfully');
				}
			} catch (error) {
				console.warn('CSS injection failed:', error);
			}
		},
		[hideBreadcrumbs, hideContentTypeField, cssInjected]
	);

	// Enhanced DOM manipulation for bulk selection
	const performAdvancedBulkSelection = useCallback(
		(doc: Document, attempt = 0) => {
			const maxAttempts = 15;
			if (attempt >= maxAttempts) {
				console.warn('Max bulk selection attempts reached');
				return;
			}

			try {
				setLoadingState((prev) => ({
					...prev,
					message: `Selecting files (attempt ${attempt + 1})...`,
					progress: (attempt / maxAttempts) * 100,
				}));

				// Wait for page readiness
				if (doc.readyState !== 'complete' || !doc.body) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 500);
					return;
				}

				// Modern SharePoint selectors (updated for latest SPO)
				const selectors = {
					fileItems: [
						'[data-automationid="DetailsRow"]',
						'[role="row"][data-selection-index]',
						'.ms-DetailsList-row',
						'[data-list-index]',
						'[data-item-index]',
					],
					checkboxes: [
						'[data-selection-toggle="true"]',
						'[role="checkbox"]',
						'.ms-Check input',
						'[data-automationid="SelectionCheckbox"]',
					],
					detailsPane: [
						'[data-automationid="SidePanelHeaderButton"]',
						'[data-automationid="InfoPaneHeaderButton"]',
						'[aria-label*="Information panel"]',
						'[aria-label*="Details panel"]',
						'button[name="Details"]',
					],
				};

				let selectedCount = 0;

				// Method 1: Enhanced filename-based selection
				for (const selector of selectors.fileItems) {
					const rows = doc.querySelectorAll(selector);
					if (rows.length === 0) continue;

					itemDetails.forEach((item) => {
						const fileName = item.fileName;
						Array.from(rows).forEach((row: Element) => {
							const rowElement = row as HTMLElement;
							const rowText = rowElement.innerText || rowElement.textContent || '';

							if (rowText.includes(fileName)) {
								// Try multiple checkbox selection methods
								for (const checkboxSelector of selectors.checkboxes) {
									const checkbox = row.querySelector(checkboxSelector) as HTMLElement;
									if (checkbox) {
										try {
											// Check if already selected
											const isSelected =
												checkbox.getAttribute('aria-checked') === 'true' ||
												checkbox.getAttribute('checked') === 'true' ||
												rowElement.getAttribute('aria-selected') === 'true';

											if (!isSelected) {
												checkbox.click();
												selectedCount++;
												console.log(`✓ Selected: ${fileName}`);
											}
											break;
										} catch (e) {
											console.warn(`Failed to select ${fileName}:`, e);
										}
									}
								}
							}
						});
					});

					if (selectedCount > 0) break; // Found working selector
				}

				// Method 2: Use SharePoint's selection API if available
				if (selectedCount === 0) {
					try {
						const spWindow = iframeRef.current?.contentWindow as any;
						if (spWindow?.__sp && spWindow.__sp.selection) {
							const itemIds = itemDetails.map((item) => item.id);
							spWindow.__sp.selection.selectItems(itemIds);
							selectedCount = itemIds.length;
							console.log('✓ Used SharePoint selection API');
						}
					} catch (e) {
						console.warn('SharePoint API selection failed:', e);
					}
				}

				// Method 3: Modern React-based selection
				if (selectedCount === 0) {
					try {
						// Look for modern selection manager
						const reactRoot = doc.querySelector('[data-reactroot]') || doc.body;
						const reactKey = Object.keys(reactRoot).find(
							(key) => key.startsWith('__reactInternalInstance') || key.startsWith('__reactFiber')
						);

						if (reactKey) {
							console.log('React root found, attempting React-based selection');
							// This would require more complex React fiber traversal
							// For now, fall back to DOM methods
						}
					} catch (e) {
						console.warn('React-based selection failed:', e);
					}
				}

				console.log(`Selection result: ${selectedCount}/${itemDetails.length} files selected`);

				// Open details pane after selection
				if (selectedCount > 0) {
					setTimeout(() => {
						for (const selector of selectors.detailsPane) {
							const button = doc.querySelector(selector) as HTMLElement;
							if (button) {
								try {
									button.click();
									console.log('✓ Opened details pane');
									setLoadingState((prev) => ({
										...prev,
										message: 'Files selected and details pane opened',
									}));
									break;
								} catch (e) {
									console.warn('Failed to open details pane:', e);
								}
							}
						}
					}, 800);
				}

				// Retry if we didn't select enough items
				if (selectedCount < itemDetails.length && attempt < maxAttempts - 1) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 1500);
				} else {
					setLoadingState((prev) => ({ ...prev, isLoading: false, message: 'Ready' }));
				}
			} catch (error) {
				console.error('Bulk selection error:', error);
				if (attempt < maxAttempts - 1) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 1000);
				} else {
					setLoadingState((prev) => ({
						...prev,
						isLoading: false,
						message: 'Selection completed with errors',
					}));
				}
			}
		},
		[itemDetails]
	);

	// Reset state when props change
	useEffect(() => {
		singleInitialLoadSeenRef.current = false;
		domNudgeAttempts.current = 0;
		setCssInjected(false);
		setError(null);
		setLoadingState({ isLoading: true, message: 'Initializing...' });

		if (changeDetectionTimer.current) {
			clearInterval(changeDetectionTimer.current);
		}
	}, [itemIds, libraryServerRelativeUrl, siteUrl]);

	// Modal visibility management
	useEffect(() => {
		if (renderMode === 'modal') {
			setModalOpen(!!isOpen);
		}
	}, [renderMode, isOpen]);

	// Enhanced initialization
	useEffect(() => {
		let disposed = false;

		const initializeEditor = async () => {
			if (!itemIds?.length) return;

			setLoadingState({ isLoading: true, message: 'Preparing editor...', progress: 10 });
			setError(null);

			const single = itemIds.length === 1;
			setMode(single ? 'single' : 'bulk');

			try {
				if (single) {
					setLoadingState({ isLoading: true, message: 'Loading edit form...', progress: 50 });

					const listId = await resolveListId(sp, libraryServerRelativeUrl);
					const url = buildSingleEditUrl(siteUrl, listId, itemIds[0], window.location.href);

					if (disposed) return;

					setTargetUrl(url);
					onDetermined?.({ mode: 'single', url, bulk: false });

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener,noreferrer');
						onOpen?.({ mode: 'single', url });
						onDismiss?.();
						return;
					} else if (renderMode === 'samepage') {
						window.location.href = url;
						onOpen?.({ mode: 'single', url });
						onDismiss?.();
						return;
					}
				} else {
					setLoadingState({ isLoading: true, message: 'Retrieving file details...', progress: 30 });

					const details = await getItemDetails(sp, libraryServerRelativeUrl, itemIds);

					if (disposed) return;

					if (details.length === 0) {
						throw new Error('Could not retrieve details for any items');
					}

					setItemDetails(details);
					setLoadingState({ isLoading: true, message: 'Building bulk edit URL...', progress: 60 });

					const url = buildBulkViewUrl(siteUrl, libraryServerRelativeUrl, details, viewId);
					setTargetUrl(url);
					onDetermined?.({ mode: 'bulk', url, bulk: true });

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener,noreferrer');
						onOpen?.({ mode: 'bulk', url });
						onDismiss?.();
						return;
					} else if (renderMode === 'samepage') {
						window.location.href = url;
						onOpen?.({ mode: 'bulk', url });
						onDismiss?.();
						return;
					}
				}

				setLoadingState({ isLoading: true, message: 'Loading interface...', progress: 90 });
			} catch (e) {
				console.error('Editor initialization failed:', e);
				setError(e instanceof Error ? e.message : 'Failed to initialize editor');
				setLoadingState({ isLoading: false, message: 'Error occurred' });
			}
		};

		initializeEditor();

		return () => {
			disposed = true;
		};
	}, [
		siteUrl,
		libraryServerRelativeUrl,
		itemIds,
		viewId,
		renderMode,
		sp,
		onDetermined,
		onOpen,
		onDismiss,
	]);

	// Enhanced change detection
	useEffect(() => {
		if (mode !== 'bulk' || !enableBulkAutoRefresh || renderMode !== 'modal') return;

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
						console.log(`✓ Changes detected in item ${id}, closing editor`);
						setModalOpen(false);
						onSaved?.();
						onDismiss?.();
						return;
					}
				}
			} catch (error) {
				console.warn('Change detection error:', error);
			}
		};

		changeDetectionTimer.current = window.setInterval(checkForChanges, 2000);
		checkForChanges();

		return () => {
			if (changeDetectionTimer.current) {
				clearInterval(changeDetectionTimer.current);
			}
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

	// Enhanced iframe load handler
	const onIframeLoad = useCallback(() => {
		const frame = iframeRef.current;
		if (!frame) return;

		console.log(`🚀 Iframe loaded: ${mode} mode, ${itemIds.length} items`);

		if (targetUrl) {
			onOpen?.({ mode, url: targetUrl });
		}

		// Inject CSS immediately
		const doc = frame.contentDocument || frame.contentWindow?.document;
		if (doc) {
			// Wait for document ready, then inject CSS
			if (doc.readyState === 'complete') {
				injectAdvancedCSS(doc);
			} else {
				doc.addEventListener('DOMContentLoaded', () => injectAdvancedCSS(doc));
			}
		}

		// Single item handling
		if (mode === 'single') {
			try {
				const href = (frame.contentWindow as any)?.location?.href as string | undefined;
				if (href) {
					if (!singleInitialLoadSeenRef.current) {
						singleInitialLoadSeenRef.current = true;
						setLoadingState({ isLoading: false, message: 'Edit form ready' });
					} else {
						const leftEditForm = !isListFormEditUrl(href);
						const returnedToHost = decodeURIComponent(href).includes(window.location.href);

						if (leftEditForm && returnedToHost) {
							console.log('✓ Single edit completed, closing');
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
			} catch (error) {
				console.warn('Single edit navigation detection failed:', error);
			}
		}

		// Bulk selection handling
		if (!disableDomNudges && mode === 'bulk' && itemDetails.length > 0 && doc) {
			// Wait for the page to render, then start selection
			setTimeout(() => {
				performAdvancedBulkSelection(doc, 0);
			}, 1200);
		} else if (mode === 'bulk') {
			setLoadingState({ isLoading: false, message: 'Bulk view ready' });
		}
	}, [
		mode,
		targetUrl,
		itemIds.length,
		itemDetails,
		disableDomNudges,
		injectAdvancedCSS,
		performAdvancedBulkSelection,
		onOpen,
		onSaved,
		onDismiss,
	]);

	// Don't render for non-modal modes
	if (renderMode === 'newtab' || renderMode === 'samepage') return null;

	// Enhanced dialog styling
	const dialogContentProps: IDialogContentProps = {
		type: DialogType.close,
		title:
			mode === 'bulk'
				? `Edit Properties - ${itemIds.length} file${itemIds.length > 1 ? 's' : ''}`
				: 'Edit File Properties',
		showCloseButton: true,
		styles: {
			content: {
				padding: '0',
				margin: '0',
			},
			header: {
				padding: '12px 20px 8px',
				borderBottom: '1px solid #edebe9',
			},
			title: {
				fontSize: '18px',
				fontWeight: '600',
			},
		},
	};

	const iframeStyle: React.CSSProperties = {
		width: '100%',
		border: 'none',
		height: autoHeightBestEffort
			? `${Math.max(600, Math.floor(window.innerHeight * 0.85))}px`
			: '85vh',
		overflow: 'hidden',
		display: loadingState.isLoading ? 'none' : 'block',
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
					main: {
						maxHeight: '95vh',
						height: 'auto',
						padding: '0',
						margin: '0',
					},
					scrollableContent: {
						padding: '0',
						margin: '0',
						overflow: 'hidden',
					},
				},
			}}
		>
			<Stack tokens={{ childrenGap: 0 }} styles={{ root: { height: '100%' } }}>
				{/* Enhanced loading indicator */}
				{loadingState.isLoading && (
					<Stack
						tokens={{ childrenGap: 12 }}
						styles={{
							root: {
								padding: '20px',
								minHeight: '200px',
								justifyContent: 'center',
								alignItems: 'center',
							},
						}}
					>
						<Spinner size={SpinnerSize.large} />
						<Text variant="medium" styles={{ root: { textAlign: 'center' } }}>
							{loadingState.message}
						</Text>
						{typeof loadingState.progress === 'number' && (
							<ProgressIndicator
								percentComplete={loadingState.progress / 100}
								description={`${Math.round(loadingState.progress)}% complete`}
								styles={{ root: { width: '300px' } }}
							/>
						)}
						{mode === 'bulk' && itemDetails.length > 0 && (
							<Text variant="small" styles={{ root: { color: '#605e5c', textAlign: 'center' } }}>
								Preparing bulk edit for: {itemDetails.map((item) => item.fileName).join(', ')}
							</Text>
						)}
					</Stack>
				)}

				{/* Error display */}
				{error && (
					<Stack styles={{ root: { padding: '16px' } }}>
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
											setLoadingState({ isLoading: true, message: 'Retrying...' });
											// Trigger re-initialization by updating a ref or state
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

				{/* Status bar for bulk operations */}
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
								fontSize: '13px',
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

				{/* Main iframe content */}
				{targetUrl && !error && (
					<div style={{ flex: 1, position: 'relative', overflow: 'hidden' }}>
						<iframe
							ref={iframeRef}
							title={mode === 'bulk' ? 'Bulk Edit Properties' : 'Edit File Properties'}
							src={targetUrl}
							style={iframeStyle}
							onLoad={onIframeLoad}
							sandbox={sandbox}
							loading="lazy"
						/>

						{/* Overlay for loading state */}
						{loadingState.isLoading && (
							<div
								style={{
									position: 'absolute',
									top: 0,
									left: 0,
									right: 0,
									bottom: 0,
									backgroundColor: 'rgba(255, 255, 255, 0.9)',
									display: 'flex',
									alignItems: 'center',
									justifyContent: 'center',
									zIndex: 1000,
								}}
							>
								<Stack tokens={{ childrenGap: 12 }} horizontalAlign="center">
									<Spinner size={SpinnerSize.large} />
									<Text variant="medium">{loadingState.message}</Text>
									{typeof loadingState.progress === 'number' && (
										<ProgressIndicator
											percentComplete={loadingState.progress / 100}
											styles={{ root: { width: '200px' } }}
										/>
									)}
								</Stack>
							</div>
						)}
					</div>
				)}

				{/* Footer with helpful information */}
				<Stack
					horizontal
					horizontalAlign="space-between"
					verticalAlign="center"
					styles={{
						root: {
							padding: '8px 16px',
							borderTop: '1px solid #edebe9',
							backgroundColor: '#fafafa',
							fontSize: '12px',
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
