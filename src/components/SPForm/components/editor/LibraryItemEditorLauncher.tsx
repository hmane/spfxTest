// src/webparts/UploadAndEdit/components/editor/LibraryItemEditorLauncher.tsx
// Final version with all fixes - TypeScript compliant
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
			[data-testid*="content-type" i]
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

/** Resolve list GUID with enhanced error handling and retry logic */
async function resolveListId(sp: SPFI, libraryServerRelativeUrl: string): Promise<string> {
	try {
		console.log('🔍 Resolving list ID for library:', libraryServerRelativeUrl);

		// Add retry logic for network issues
		let lastError: Error | null = null;
		for (let attempt = 0; attempt < 3; attempt++) {
			try {
				const list = await sp.web.getList(libraryServerRelativeUrl).select('Id')();
				const listId = list?.Id as string;

				if (!listId) {
					throw new Error('List ID not found in response');
				}

				console.log('✅ List ID resolved successfully:', listId);
				return listId;
			} catch (error) {
				lastError = error as Error;
				console.warn(`Attempt ${attempt + 1} failed:`, error);

				if (attempt < 2) {
					// Wait before retry
					await new Promise((resolve) => setTimeout(resolve, 1000 * (attempt + 1)));
				}
			}
		}

		throw lastError || new Error('Failed to resolve list ID after 3 attempts');
	} catch (error) {
		console.error('❌ List ID resolution failed:', error);

		// Provide helpful error messages based on common issues
		if (error instanceof Error) {
			if (error.message.includes('404') || error.message.includes('not found')) {
				throw new Error(
					`Library not found: ${libraryServerRelativeUrl}. Please check the library path and permissions.`
				);
			} else if (error.message.includes('403') || error.message.includes('Forbidden')) {
				throw new Error('Access denied. You may not have permission to access this library.');
			} else if (error.message.includes('timeout')) {
				throw new Error('Request timeout. Please check your network connection and try again.');
			}
		}

		throw new Error(
			`Unable to access library: ${error instanceof Error ? error.message : 'Unknown error'}`
		);
	}
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

/**
 * Fixed bulk view URL construction - prevents duplicate site paths
 * and uses modern SharePoint selection approach
 */
function buildBulkViewUrl(
	siteUrl: string,
	libraryServerRelativeUrl: string,
	itemDetails: ItemInfo[],
	viewId?: string
): string {
	// Clean the siteUrl - remove trailing slash
	const cleanSiteUrl = siteUrl.replace(/\/$/, '');

	// IMPORTANT: libraryServerRelativeUrl already contains the full path from root
	// Don't concatenate with siteUrl again to avoid duplication!
	let url = `${cleanSiteUrl}${libraryServerRelativeUrl}`;

	console.log('🔗 Building bulk URL:', {
		siteUrl: cleanSiteUrl,
		libraryPath: libraryServerRelativeUrl,
		finalBaseUrl: url,
		itemCount: itemDetails.length,
	});

	// For modern SharePoint, use proper parameters
	const params = new URLSearchParams();

	// Add view if specified (modern format)
	if (viewId) {
		const cleanViewId = viewId.replace(/[{}]/g, '');
		params.set('viewid', `{${cleanViewId}}`);
	}

	// Modern SharePoint bulk selection approach
	if (itemDetails.length > 0) {
		// Use item IDs for better compatibility
		const itemIds = itemDetails.map((item) => item.id.toString());

		// Method 1: Filter to show only our uploaded files (easier to select)
		const fileNames = itemDetails.map((item) => item.fileName).join(',');
		params.set('FilterField1', 'FileLeafRef');
		params.set('FilterValue1', fileNames);
		params.set('FilterType1', 'Text');

		// Method 2: Force modern experience
		params.set('env', 'WebView');
		params.set('OR', 'Teams-HL'); // Forces modern experience

		// Method 3: Add selection hints
		params.set('selectedItems', itemIds.join(','));
	}

	// Construct final URL
	if (params.toString()) {
		url += `?${params.toString()}`;
	}

	console.log('🎯 Final bulk URL:', url);
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
	const changeDetectionTimer = useRef<number>();
	const loadingTimeoutRef = useRef<number>();

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

	// Enhanced DOM manipulation for modern bulk selection
	const performAdvancedBulkSelection = useCallback(
		(doc: Document, attempt = 0) => {
			const maxAttempts = 20; // Increased for modern lists that load slower
			if (attempt >= maxAttempts) {
				console.warn('⚠️ Max modern bulk selection attempts reached');
				setLoadingState((prev) => ({
					...prev,
					isLoading: false,
					message: 'Selection completed (some items may not be selected)',
				}));
				return;
			}

			try {
				setLoadingState((prev) => ({
					...prev,
					message: `Selecting files for bulk edit (${attempt + 1}/${maxAttempts})...`,
					progress: (attempt / maxAttempts) * 100,
				}));

				console.log(`🔄 Modern selection attempt ${attempt + 1}:`, {
					readyState: doc.readyState,
					itemCount: itemDetails.length,
					bodyExists: !!doc.body,
					url: doc.location?.href,
				});

				// Wait for page readiness
				if (doc.readyState !== 'complete' || !doc.body) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 800);
					return;
				}

				// Modern SharePoint Online selectors (updated for 2024/2025)
				const modernSelectors = {
					// File list items in modern view
					listItems: [
						'[data-automationid="DetailsRow"]',
						'[role="row"][data-selection-index]',
						'[data-list-index]',
						'div[data-automationid="DetailsRow"]',
						'.ms-DetailsRow',
						'[role="gridcell"]',
						'.od-ItemContent-file',
					],
					// Selection checkboxes
					checkboxes: [
						'[data-selection-toggle="true"]',
						'button[data-selection-toggle="true"]',
						'[role="checkbox"]',
						'[data-automationid="DetailsRowCheck"]',
						'.ms-Check input[type="checkbox"]',
						'[aria-label*="Select row"]',
						'[data-automationid="SelectionCheckbox"]',
					],
					// Details pane / property panel
					detailsPane: [
						'[data-automationid="PropertyPaneButton"]',
						'[data-automationid="InfoButton"]',
						'[aria-label*="Information panel"]',
						'[aria-label*="Details panel"]',
						'[aria-label*="Properties"]',
						'button[name="Details"]',
						'button[name="Properties"]',
						'[data-automationid="SidePanelHeaderButton"]',
					],
					// Command bar for bulk operations
					commandBar: ['[data-automationid="CommandBar"]', '.ms-CommandBar', '[role="menubar"]'],
				};

				let selectedCount = 0;
				let foundItems = 0;

				// Method 1: Modern list row selection with improved matching
				for (const selector of modernSelectors.listItems) {
					const rows = doc.querySelectorAll(selector);
					if (rows.length === 0) continue;

					console.log(`📋 Found ${rows.length} list items with selector: ${selector}`);
					foundItems = rows.length;

					// Try to select items by filename matching
					itemDetails.forEach((item) => {
						const fileName = item.fileName;
						const fileNameWithoutExt = fileName.replace(/\.[^/.]+$/, '');
						let itemSelected = false;

						Array.from(rows).forEach((row: Element) => {
							if (itemSelected) return; // Skip if already selected this item

							const rowElement = row as HTMLElement;
							const rowText = rowElement.innerText || rowElement.textContent || '';

							// Enhanced filename matching
							const nameMatches =
								rowText.includes(fileName) ||
								rowText.includes(fileNameWithoutExt) ||
								rowElement.querySelector(`[title*="${fileName}"]`) ||
								rowElement.querySelector(`[aria-label*="${fileName}"]`) ||
								rowElement.querySelector(`[href*="${encodeURIComponent(fileName)}"]`) ||
								// Check data attributes
								rowElement.getAttribute('data-item-id') === item.id.toString() ||
								rowElement.getAttribute('data-unique-id') === item.uniqueId;

							if (nameMatches) {
								console.log(`🎯 Found matching row for: ${fileName}`);

								// Try different checkbox selection methods
								for (const checkboxSelector of modernSelectors.checkboxes) {
									const checkbox = row.querySelector(checkboxSelector) as HTMLElement;
									if (checkbox) {
										try {
											// Check current selection state
											const isAlreadySelected =
												checkbox.getAttribute('aria-checked') === 'true' ||
												checkbox.getAttribute('checked') === 'true' ||
												rowElement.getAttribute('aria-selected') === 'true' ||
												rowElement.classList.contains('is-selected') ||
												rowElement.classList.contains('ms-DetailsRow--selected');

											if (!isAlreadySelected) {
												// Try multiple click methods for modern UI
												checkbox.click();

												// Dispatch proper events
												const events = ['mousedown', 'mouseup', 'click'];
												events.forEach((eventType) => {
													const event = new MouseEvent(eventType, {
														bubbles: true,
														cancelable: true,
														view: doc.defaultView || undefined,
													});
													checkbox.dispatchEvent(event);
												});

												selectedCount++;
												itemSelected = true;
												console.log(`✅ Selected: ${fileName} (method: checkbox)`);
											} else {
												console.log(`ℹ️ Already selected: ${fileName}`);
												selectedCount++;
												itemSelected = true;
											}
											break;
										} catch (e) {
											console.warn(`❌ Checkbox click failed for ${fileName}:`, e);
										}
									}
								}

								// If checkbox didn't work, try clicking the row itself
								if (!itemSelected) {
									try {
										// Find the clickable area within the row
										const clickableArea =
											rowElement.querySelector('[data-selection-index]') ||
											rowElement.querySelector('[role="gridcell"]:first-child') ||
											rowElement;

										(clickableArea as HTMLElement).click();

										// Also try focus + enter key
										setTimeout(() => {
											try {
												(clickableArea as HTMLElement).focus();
												const enterEvent = new KeyboardEvent('keydown', {
													key: 'Enter',
													bubbles: true,
													cancelable: true,
												});
												(clickableArea as HTMLElement).dispatchEvent(enterEvent);
											} catch (e) {
												console.warn('Focus+Enter failed:', e);
											}
										}, 100);

										selectedCount++;
										itemSelected = true;
										console.log(`✅ Selected: ${fileName} (method: row click)`);
									} catch (e) {
										console.warn(`❌ Row click failed for ${fileName}:`, e);
									}
								}
							}
						});
					});

					if (selectedCount > 0) break; // Found working method
				}

				console.log(`📊 Selection progress: ${selectedCount}/${itemDetails.length} items selected`);

				// Method 2: Try bulk select all if individual selection failed
				if (selectedCount === 0 && attempt > 8) {
					try {
						console.log('🔄 Trying bulk select all approach...');
						const selectAllCheckbox = doc.querySelector(
							'[data-automationid="SelectAllCheckbox"]'
						) as HTMLElement;
						if (selectAllCheckbox) {
							selectAllCheckbox.click();
							selectedCount = itemDetails.length; // Assume success
							console.log('✅ Used select-all checkbox');
						}
					} catch (e) {
						console.warn('Select-all failed:', e);
					}
				}

				// Method 3: Force details pane open if items are selected
				if (selectedCount > 0) {
					setTimeout(() => {
						let panelOpened = false;

						for (const selector of modernSelectors.detailsPane) {
							if (panelOpened) break;

							const button = doc.querySelector(selector) as HTMLElement;
							if (button) {
								try {
									button.click();
									panelOpened = true;
									console.log(`✅ Opened details panel with: ${selector}`);
									break;
								} catch (e) {
									console.warn(`❌ Failed to open panel with ${selector}:`, e);
								}
							}
						}

						// If no specific details button found, try generic approaches
						if (!panelOpened) {
							setTimeout(() => {
								// Try right-click context menu approach
								const firstSelectedRow = doc.querySelector(
									'.ms-DetailsRow--selected, [aria-selected="true"]'
								) as HTMLElement;
								if (firstSelectedRow) {
									try {
										const contextMenuEvent = new MouseEvent('contextmenu', {
											bubbles: true,
											cancelable: true,
											clientX: 100,
											clientY: 100,
										});
										firstSelectedRow.dispatchEvent(contextMenuEvent);
										console.log('✅ Triggered context menu');
									} catch (e) {
										console.warn('Context menu failed:', e);
									}
								}
							}, 500);
						}
					}, 1200); // Give time for selection to register
				}

				// Continue trying if we haven't selected enough items
				if (selectedCount < itemDetails.length && attempt < maxAttempts - 1) {
					const delay = Math.min(3000, 1000 + attempt * 300); // Increasing delay for modern lists
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), delay);
				} else {
					// Final status
					const successRate = (selectedCount / itemDetails.length) * 100;
					const finalMessage =
						selectedCount === itemDetails.length
							? `✅ All ${selectedCount} files selected successfully!`
							: `Selection complete: ${selectedCount}/${itemDetails.length} files (${Math.round(
									successRate
							  )}%)`;

					setLoadingState((prev) => ({
						...prev,
						isLoading: false,
						message: finalMessage,
					}));

					if (selectedCount === 0 && foundItems > 0) {
						console.warn('⚠️ Found list items but could not select any. This might be due to:');
						console.warn('   - Permissions restrictions');
						console.warn('   - Modern list UI changes');
						console.warn('   - JavaScript restrictions');
						console.warn('   - Network delays');
					}
				}
			} catch (error) {
				console.error('❌ Modern bulk selection error:', error);
				if (attempt < maxAttempts - 1) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 2000);
				} else {
					setLoadingState((prev) => ({
						...prev,
						isLoading: false,
						message:
							'⚠️ Auto-selection failed - please select files manually and use the details pane',
					}));
				}
			}
		},
		[itemDetails, setLoadingState]
	);

	// Cleanup effect for timers and resources
	useEffect(() => {
		return () => {
			// Cleanup all timers on unmount
			if (changeDetectionTimer.current) {
				clearInterval(changeDetectionTimer.current);
			}
			if (loadingTimeoutRef.current) {
				clearTimeout(loadingTimeoutRef.current);
			}
		};
	}, []);

	// Reset state when props change
	useEffect(() => {
		singleInitialLoadSeenRef.current = false;
		setCssInjected(false);
		setError(null);
		setLoadingState({ isLoading: true, message: 'Initializing...' });

		// Clear any existing timers
		if (changeDetectionTimer.current) {
			clearInterval(changeDetectionTimer.current);
		}
		if (loadingTimeoutRef.current) {
			clearTimeout(loadingTimeoutRef.current);
		}

		// Set a safety timeout to prevent indefinite loading
		// Set a safety timeout to prevent indefinite loading
		loadingTimeoutRef.current = window.setTimeout(() => {
			console.warn('⚠️ Loading timeout reached, clearing loading state');
			setLoadingState((prev) =>
				prev.isLoading
					? {
							isLoading: false,
							message: 'Loading timeout - please try again',
					  }
					: prev
			);
		}, 30000); // 30 second timeout
	}, [itemIds, libraryServerRelativeUrl, siteUrl]);

	// Modal visibility management
	useEffect(() => {
		if (renderMode === 'modal') {
			setModalOpen(!!isOpen);
		}
	}, [renderMode, isOpen]);

	// Enhanced initialization with better error handling and progress tracking
	useEffect(() => {
		let disposed = false;

		const initializeEditor = async () => {
			if (!itemIds?.length) return;

			try {
				setLoadingState({ isLoading: true, message: 'Initializing editor...', progress: 0 });
				setError(null);

				const single = itemIds.length === 1;
				setMode(single ? 'single' : 'bulk');

				if (single) {
					// Single item edit with detailed progress tracking
					try {
						setLoadingState({
							isLoading: true,
							message: 'Resolving list information...',
							progress: 20,
						});

						console.log('🔍 Resolving list ID for:', libraryServerRelativeUrl);
						const listId = await Promise.race([
							resolveListId(sp, libraryServerRelativeUrl),
							new Promise<never>((_, reject) =>
								setTimeout(() => reject(new Error('List ID resolution timeout')), 15000)
							),
						]);

						if (disposed) return;
						console.log('✅ List ID resolved:', listId);

						setLoadingState({
							isLoading: true,
							message: 'Building edit form URL...',
							progress: 60,
						});

						const url = buildSingleEditUrl(siteUrl, listId, itemIds[0], window.location.href);
						console.log('🔗 Single edit URL:', url);

						if (disposed) return;

						setTargetUrl(url);
						setLoadingState({ isLoading: true, message: 'Loading edit form...', progress: 90 });

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

						// For modal mode, the iframe will handle the final loading
						setLoadingState({ isLoading: true, message: 'Ready to load form...', progress: 100 });
					} catch (singleError) {
						console.error('❌ Single edit initialization failed:', singleError);
						throw new Error(
							`Failed to prepare edit form: ${
								singleError instanceof Error ? singleError.message : 'Unknown error'
							}`
						);
					}
				} else {
					// Bulk edit with detailed progress tracking
					try {
						setLoadingState({
							isLoading: true,
							message: 'Retrieving file details...',
							progress: 20,
						});

						console.log('📋 Getting details for', itemIds.length, 'items');
						const details = await Promise.race([
							getItemDetails(sp, libraryServerRelativeUrl, itemIds),
							new Promise<never>((_, reject) =>
								setTimeout(() => reject(new Error('Item details timeout')), 20000)
							),
						]);

						if (disposed) return;
						console.log('✅ Retrieved details for', details.length, 'items');

						if (details.length === 0) {
							throw new Error(
								'Could not retrieve details for any files. Please check permissions and try again.'
							);
						}

						setItemDetails(details);
						setLoadingState({
							isLoading: true,
							message: 'Building bulk edit URL...',
							progress: 60,
						});

						const url = buildBulkViewUrl(siteUrl, libraryServerRelativeUrl, details, viewId);
						console.log('🔗 Bulk edit URL:', url);

						if (disposed) return;

						setTargetUrl(url);
						setLoadingState({ isLoading: true, message: 'Loading library view...', progress: 90 });

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

						// For modal mode, the iframe will handle the final loading
						setLoadingState({
							isLoading: true,
							message: 'Ready to load library...',
							progress: 100,
						});
					} catch (bulkError) {
						console.error('❌ Bulk edit initialization failed:', bulkError);
						throw new Error(
							`Failed to prepare bulk edit: ${
								bulkError instanceof Error ? bulkError.message : 'Unknown error'
							}`
						);
					}
				}
			} catch (error) {
				console.error('❌ Editor initialization failed:', error);
				if (!disposed) {
					const errorMessage =
						error instanceof Error ? error.message : 'Failed to initialize editor';
					setError(errorMessage);
					setLoadingState({ isLoading: false, message: 'Initialization failed' });
				}
			}
		};

		// Add a small delay to prevent immediate execution issues
		const initTimer = setTimeout(() => {
			void initializeEditor();
		}, 100);

		return () => {
			disposed = true;
			clearTimeout(initTimer);
			// Clear loading timeout when component unmounts or dependencies change
			if (loadingTimeoutRef.current) {
				clearTimeout(loadingTimeoutRef.current);
			}
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

	// Enhanced auto-refresh with better change detection
	useEffect(() => {
		if (mode !== 'bulk' || !enableBulkAutoRefresh || renderMode !== 'modal') return;

		const idsToWatch = bulkWatchAllItems ? itemIds : [itemIds[0]];
		const originalModified: Record<number, string> = {};
		let initialized = false;

		const checkForChanges = async () => {
			try {
				if (!initialized) {
					// Store initial Modified timestamps
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

				// Check for changes
				for (const id of idsToWatch) {
					const item: any = await sp.web
						.getList(libraryServerRelativeUrl)
						.items.getById(id)
						.select('Modified')();

					if (item?.Modified && originalModified[id] && item.Modified !== originalModified[id]) {
						console.log(`✅ Detected change in item ${id}, closing editor`);
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

		changeDetectionTimer.current = window.setInterval(checkForChanges, 3000); // Check every 3 seconds
		void checkForChanges(); // Initial check

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

	// Enhanced iframe load handler with proper loading state management
	const onIframeLoad = useCallback(() => {
		const frame = iframeRef.current;
		if (!frame) return;

		console.log(`🚀 Iframe loaded: ${mode} mode, ${itemIds.length} items`);

		// Clear the loading state immediately when iframe loads
		if (mode === 'single') {
			setLoadingState({ isLoading: false, message: 'Edit form ready' });
		} else {
			setLoadingState({ isLoading: true, message: 'Preparing bulk selection...', progress: 0 });
		}

		if (targetUrl) {
			onOpen?.({ mode, url: targetUrl });
		}

		// Inject CSS immediately when iframe loads
		const doc = frame.contentDocument || frame.contentWindow?.document;
		if (doc) {
			// Wait for document ready, then inject CSS
			const injectCSS = () => {
				try {
					injectAdvancedCSS(doc);
				} catch (error) {
					console.warn('CSS injection failed:', error);
				}
			};

			if (doc.readyState === 'complete') {
				injectCSS();
			} else {
				doc.addEventListener('DOMContentLoaded', injectCSS);
				// Fallback timeout
				setTimeout(injectCSS, 1000);
			}
		}

		// Single item handling
		if (mode === 'single') {
			try {
				const href = (frame.contentWindow as any)?.location?.href as string | undefined;
				if (href) {
					if (!singleInitialLoadSeenRef.current) {
						singleInitialLoadSeenRef.current = true;
						console.log('✅ Single edit form loaded successfully');
					} else {
						const leftEditForm = !isListFormEditUrl(href);
						const returnedToHost = decodeURIComponent(href).includes(window.location.href);

						if (leftEditForm && returnedToHost) {
							console.log('✅ Single edit completed, closing');
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

		// Enhanced bulk selection handling for modern lists
		if (!disableDomNudges && mode === 'bulk' && itemDetails.length > 0 && doc) {
			// Wait longer for modern SharePoint lists to fully load
			setTimeout(() => {
				console.log('🚀 Starting modern bulk selection process...');
				performAdvancedBulkSelection(doc, 0);
			}, 2000); // Increased delay for modern lists
		} else if (mode === 'bulk') {
			setLoadingState({
				isLoading: false,
				message: 'Bulk view ready - please select files manually',
			});
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
											// Force re-initialization by updating state
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

				{/* Main iframe content with error recovery */}
				{targetUrl && !error && (
					<div style={{ flex: 1, position: 'relative', overflow: 'hidden' }}>
						<iframe
							ref={iframeRef}
							title={mode === 'bulk' ? 'Bulk Edit Properties' : 'Edit File Properties'}
							src={targetUrl}
							style={iframeStyle}
							onLoad={onIframeLoad}
							onError={(e) => {
								console.error('Iframe load error:', e);
								setError('Failed to load the edit form. Please try again or refresh the page.');
								setLoadingState({ isLoading: false, message: 'Load failed' });
							}}
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
