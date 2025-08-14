// src/webparts/UploadAndEdit/components/editor/LibraryItemEditorLauncher.tsx
// POLISHED FINAL VERSION - All TypeScript and ESLint issues fixed
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
	serverRelativeUrl: string;
	listItemId: number;
}

interface LoadingState {
	isLoading: boolean;
	message: string;
	progress?: number;
}

interface InteractionState {
	isInteracting: boolean;
	message: string;
	progress?: number;
}

interface InitializationResult {
	url: string;
	mode: 'single' | 'bulk';
	details: ItemInfo[];
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

const createAdvancedCSS = (hideBreadcrumbs: boolean, hideContentTypeField: boolean): string => {
	const css: string[] = [];

	// Base styles
	css.push(`
		body {
			margin: 0 !important;
			padding: 0 !important;
			overflow-x: hidden !important;
		}
		.ms-Dialog-main {
			padding: 0 !important;
		}
	`);

	// Hide breadcrumbs
	if (hideBreadcrumbs) {
		css.push(`
			.od-TopNav, .od-TopBar, .ms-CommandBar, .ms-Breadcrumb, .ms-Nav,
			[data-automationid="Breadcrumb"], .sp-appBar, #spPageChromeAppDiv,
			.spSiteHeader, #SuiteNavWrapper, #s4-titlerow, #s4-ribbonrow {
				display: none !important;
				height: 0 !important;
				overflow: hidden !important;
				visibility: hidden !important;
			}
		`);
	}

	// Hide content type field
	if (hideContentTypeField) {
		css.push(`
			div[data-field="ContentType"],
			div[aria-label*="Content type" i],
			[data-automationid="ContentTypeSelector"] {
				display: none !important;
				height: 0 !important;
				overflow: hidden !important;
				visibility: hidden !important;
			}
		`);
	}

	return css.join('\n');
};

async function resolveListId(sp: SPFI, libraryServerRelativeUrl: string): Promise<string> {
	try {
		const list = await sp.web.getList(libraryServerRelativeUrl).select('Id')();
		const listId = list?.Id as string;
		if (!listId) {
			throw new Error('List ID not found in response');
		}
		return listId;
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : 'Unknown error';
		throw new Error(`Failed to resolve list ID: ${errorMessage}`);
	}
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
				id: item.Id as number,
				uniqueId: (item.UniqueId || item.GUID) as string,
				fileName: item.FileLeafRef as string,
				modified: item.Modified as string,
				serverRelativeUrl: item.FileRef as string,
				listItemId: item.Id as number,
			};
		})
	);

	const successfulResults: ItemInfo[] = [];
	results.forEach((result) => {
		if (result.status === 'fulfilled') {
			successfulResults.push(result.value);
		} else {
			console.warn('Failed to get item details:', result.reason);
		}
	});

	return successfulResults;
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
		const itemIdsStr = itemDetails.map((item) => item.id.toString()).join(',');
		const fileNames = itemDetails.map((item) => item.fileName).join(',');
		params.set('FilterField1', 'FileLeafRef');
		params.set('FilterValue1', fileNames);
		params.set('FilterType1', 'Text');
		params.set('env', 'WebView');
		params.set('selectedItems', itemIdsStr);
	}

	if (params.toString()) {
		url += `?${params.toString()}`;
	}

	return url;
}

function isListFormEditUrl(href: string): boolean {
	try {
		const u = new URL(href, window.location.origin);
		return (
			u.pathname.toLowerCase().includes('/_layouts/15/listform.aspx') &&
			u.searchParams.get('PageType') === '6'
		);
	} catch {
		return false;
	}
}

// ============================================================================
// MAIN COMPONENT
// ============================================================================

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

	// ============================================================================
	// STATE MANAGEMENT
	// ============================================================================

	const [targetUrl, setTargetUrl] = useState<string>('');
	const [mode, setMode] = useState<'single' | 'bulk'>('bulk');
	const [modalOpen, setModalOpen] = useState<boolean>(isOpen && renderMode === 'modal');
	const [itemDetails, setItemDetails] = useState<ItemInfo[]>([]);
	const [loadingState, setLoadingState] = useState<LoadingState>({
		isLoading: true,
		message: 'Initializing...',
	});
	const [error, setError] = useState<string | null>(null);
	const [interactionState, setInteractionState] = useState<InteractionState>({
		isInteracting: false,
		message: '',
		progress: 0,
	});

	// ============================================================================
	// REFS AND MEMOIZED VALUES
	// ============================================================================

	const sp = useMemo(() => spfi(siteUrl).using(PnP_SPFX(spfxContext)), [siteUrl, spfxContext]);
	const iframeRef = useRef<HTMLIFrameElement | null>(null);
	const singleInitialLoadSeenRef = useRef(false);
	const changeDetectionTimer = useRef<number>();

	const stableKey = useMemo(
		() => `${siteUrl}|${libraryServerRelativeUrl}|${itemIds.join(',')}|${renderMode}`,
		[siteUrl, libraryServerRelativeUrl, itemIds, renderMode]
	);

	// ============================================================================
	// CALLBACKS
	// ============================================================================

	const injectAdvancedCSS = useCallback(
		(doc: Document) => {
			try {
				const styleId = 'enhanced-launcher-styles';
				if (doc.getElementById(styleId)) {
					return;
				}

				const style = doc.createElement('style');
				style.id = styleId;
				style.type = 'text/css';
				style.appendChild(
					doc.createTextNode(createAdvancedCSS(hideBreadcrumbs, hideContentTypeField))
				);

				if (doc.head) {
					doc.head.appendChild(style);
				}
			} catch (error) {
				console.warn('CSS injection failed:', error);
			}
		},
		[hideBreadcrumbs, hideContentTypeField]
	);

	const performAdvancedBulkSelection = useCallback(
		(doc: Document, attempt = 0) => {
			const maxAttempts = 15;
			if (attempt >= maxAttempts) {
				setInteractionState({ isInteracting: false, message: 'Selection complete.' });
				return;
			}

			setInteractionState({
				isInteracting: true,
				message: `Selecting files (${attempt + 1}/${maxAttempts})...`,
				progress: (attempt / maxAttempts) * 100,
			});

			try {
				const modernSelectors = {
					listItems: ['[data-automationid="DetailsRow"]', '[role="row"][data-selection-index]'],
					checkboxes: ['[data-selection-toggle="true"]', '[role="checkbox"]'],
					detailsPane: [
						'[data-automationid="PropertyPaneButton"]',
						'[data-automationid="InfoButton"]',
						'button[name="Details"]',
					],
				};

				let selectedCount = 0;

				for (const selector of modernSelectors.listItems) {
					const rows = doc.querySelectorAll(selector);
					if (rows.length === 0) continue;

					itemDetails.forEach((item) => {
						Array.from(rows).forEach((row: Element) => {
							const rowElement = row as HTMLElement;
							const rowText = rowElement.innerText || rowElement.textContent || '';

							if (rowText.includes(item.fileName)) {
								const checkbox = row.querySelector(
									modernSelectors.checkboxes.join(',')
								) as HTMLElement;
								if (checkbox && checkbox.getAttribute('aria-checked') !== 'true') {
									checkbox.click();
									selectedCount++;
								}
							}
						});
					});

					if (selectedCount > 0) break;
				}

				if (selectedCount > 0) {
					setTimeout(() => {
						for (const selector of modernSelectors.detailsPane) {
							const button = doc.querySelector(selector) as HTMLElement;
							if (button) {
								button.click();
								break;
							}
						}
						setInteractionState({ isInteracting: false, message: 'Selection complete' });
					}, 1200);
				} else if (attempt < maxAttempts - 1) {
					setTimeout(() => performAdvancedBulkSelection(doc, attempt + 1), 800);
				} else {
					setInteractionState({ isInteracting: false, message: 'Auto-selection failed' });
				}
			} catch (error) {
				console.error('Bulk selection error:', error);
				setInteractionState({ isInteracting: false, message: 'Selection error' });
			}
		},
		[itemDetails]
	);

	const onIframeLoad = useCallback(() => {
		const frame = iframeRef.current;
		if (!frame) return;

		if (onOpen) {
			onOpen({ mode, url: targetUrl });
		}

		const doc = frame.contentDocument || frame.contentWindow?.document;
		if (doc) {
			injectAdvancedCSS(doc);
		}

		// Handle single edit completion detection
		if (mode === 'single') {
			try {
				const href = frame.contentWindow?.location.href;
				if (href) {
					if (!singleInitialLoadSeenRef.current) {
						singleInitialLoadSeenRef.current = true;
					} else if (
						!isListFormEditUrl(href) &&
						decodeURIComponent(href).includes(window.location.href)
					) {
						setModalOpen(false);
						if (onSaved) onSaved();
						if (onDismiss) onDismiss();
					}
				}
			} catch (error) {
				console.warn('Single edit navigation detection failed:', error);
			}
		}

		// Handle bulk selection
		if (!disableDomNudges && mode === 'bulk' && itemDetails.length > 0 && doc) {
			setTimeout(() => performAdvancedBulkSelection(doc), 2000);
		}
	}, [
		mode,
		targetUrl,
		itemDetails,
		disableDomNudges,
		injectAdvancedCSS,
		performAdvancedBulkSelection,
		onOpen,
		onSaved,
		onDismiss,
	]);

	// ============================================================================
	// EFFECTS
	// ============================================================================

	// Modal visibility management
	useEffect(() => {
		if (renderMode === 'modal') {
			setModalOpen(!!isOpen);
		}
	}, [renderMode, isOpen]);

	// Main initialization effect
	useEffect(() => {
		let isCancelled = false;

		const fetchInitializationData = async (): Promise<InitializationResult> => {
			const currentMode: 'single' | 'bulk' = itemIds.length === 1 ? 'single' : 'bulk';
			let url = '';
			let fetchedDetails: ItemInfo[] = [];

			if (currentMode === 'single') {
				const listId = await resolveListId(sp, libraryServerRelativeUrl);
				url = buildSingleEditUrl(siteUrl, listId, itemIds[0], window.location.href);
			} else {
				fetchedDetails = await getItemDetails(sp, libraryServerRelativeUrl, itemIds);
				if (fetchedDetails.length === 0) {
					throw new Error('Could not retrieve details for any files.');
				}
				url = buildBulkViewUrl(siteUrl, libraryServerRelativeUrl, fetchedDetails, viewId);
			}

			return { url, mode: currentMode, details: fetchedDetails };
		};

		// Reset state
		setError(null);
		setTargetUrl('');
		setItemDetails([]);
		setInteractionState({ isInteracting: false, message: '' });
		setLoadingState({ isLoading: true, message: 'Initializing...' });
		singleInitialLoadSeenRef.current = false;

		// Validate input
		if (!itemIds || itemIds.length === 0) {
			setLoadingState({ isLoading: false, message: 'No items selected.' });
			return;
		}

		// Set loading timeout
		const loadingTimeout = setTimeout(() => {
			if (!isCancelled) {
				setError('Initialization timed out. Please check permissions and refresh.');
				setLoadingState({ isLoading: false, message: 'Timeout' });
			}
		}, 30000);

		// Execute initialization
		fetchInitializationData()
			.then((result) => {
				if (isCancelled) return;

				const { url, mode: resultMode, details } = result;

				if (onDetermined) {
					onDetermined({ mode: resultMode, url, bulk: resultMode === 'bulk' });
				}

				// Handle non-modal modes
				if (renderMode === 'newtab' || renderMode === 'samepage') {
					if (onOpen) {
						onOpen({ mode: resultMode, url });
					}

					if (renderMode === 'newtab') {
						window.open(url, '_blank', 'noopener,noreferrer');
					} else {
						window.location.href = url;
					}

					if (onDismiss) {
						onDismiss();
					}
					return;
				}

				// Set modal state
				setMode(resultMode);
				setItemDetails(details);
				setTargetUrl(url);
				setLoadingState({ isLoading: false, message: 'Loading form...' });
			})
			.catch((err) => {
				if (isCancelled) return;

				const errorMessage = err instanceof Error ? err.message : 'An unknown error occurred.';
				setError(errorMessage);
				setLoadingState({ isLoading: false, message: 'Initialization failed' });
			});

		return () => {
			isCancelled = true;
			clearTimeout(loadingTimeout);
			if (changeDetectionTimer.current) {
				clearInterval(changeDetectionTimer.current);
			}
		};
	}, [
		stableKey,
		sp,
		viewId,
		onDetermined,
		onOpen,
		onDismiss,
		siteUrl,
		libraryServerRelativeUrl,
		itemIds,
		renderMode,
	]);

	// Auto-refresh effect for bulk mode
	useEffect(() => {
		if (mode !== 'bulk' || !enableBulkAutoRefresh || !targetUrl) {
			return;
		}

		const originalModified: Record<number, string> = {};
		let isInitialized = false;

		const checkForChanges = async (): Promise<void> => {
			try {
				const idsToWatch = bulkWatchAllItems ? itemIds : [itemIds[0]];

				if (!isInitialized) {
					const items = await getItemDetails(sp, libraryServerRelativeUrl, idsToWatch);
					items.forEach((item) => {
						originalModified[item.id] = item.modified;
					});
					isInitialized = true;
					return;
				}

				const currentItems = await getItemDetails(sp, libraryServerRelativeUrl, idsToWatch);
				for (const item of currentItems) {
					if (originalModified[item.id] && item.modified !== originalModified[item.id]) {
						setModalOpen(false);
						if (onSaved) onSaved();
						if (onDismiss) onDismiss();
						return;
					}
				}
			} catch (error) {
				console.warn('Change detection error:', error);
				if (changeDetectionTimer.current) {
					clearInterval(changeDetectionTimer.current);
				}
			}
		};

		changeDetectionTimer.current = window.setInterval(() => {
			void checkForChanges();
		}, 3500);

		void checkForChanges();

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
		targetUrl,
		sp,
		libraryServerRelativeUrl,
		onSaved,
		onDismiss,
	]);

	// ============================================================================
	// RENDER
	// ============================================================================

	// Don't render for non-modal modes
	if (renderMode !== 'modal') {
		return null;
	}

	const dialogContentProps: IDialogContentProps = {
		type: DialogType.close,
		title: mode === 'bulk' ? `Edit Properties - ${itemIds.length} file(s)` : 'Edit File Properties',
		showCloseButton: true,
	};

	const iframeStyle: React.CSSProperties = {
		width: '100%',
		border: 'none',
		height: autoHeightBestEffort
			? `${Math.max(600, Math.floor(window.innerHeight * 0.85))}px`
			: '85vh',
		visibility: interactionState.isInteracting ? 'hidden' : 'visible',
	};

	const sandbox = `allow-scripts allow-same-origin allow-forms allow-popups allow-downloads allow-modals allow-presentation${
		sandboxExtra ? ` ${sandboxExtra}` : ''
	}`;

	return (
		<Dialog
			hidden={!modalOpen}
			onDismiss={() => {
				setModalOpen(false);
				if (onDismiss) onDismiss();
			}}
			dialogContentProps={dialogContentProps}
			minWidth="75%"
			maxWidth="98%"
			modalProps={{
				isBlocking: true,
				styles: {
					main: { maxHeight: '95vh', height: 'auto' },
				},
			}}
		>
			<Stack styles={{ root: { height: '100%', overflow: 'hidden' } }}>
				{/* Loading indicator */}
				{loadingState.isLoading && (
					<Stack
						tokens={{ childrenGap: 12 }}
						styles={{ root: { padding: '40px', alignItems: 'center' } }}
					>
						<Spinner size={SpinnerSize.large} />
						<Text variant="medium">{loadingState.message}</Text>
						{typeof loadingState.progress === 'number' && (
							<ProgressIndicator
								percentComplete={loadingState.progress / 100}
								styles={{ root: { width: '300px' } }}
							/>
						)}
					</Stack>
				)}

				{/* Error display */}
				{error && (
					<MessageBar
						messageBarType={MessageBarType.error}
						onDismiss={() => setError(null)}
						actions={
							<IconButton
								iconProps={{ iconName: 'Refresh' }}
								title="Retry"
								ariaLabel="Retry"
								onClick={() => window.location.reload()}
							/>
						}
					>
						<strong>Error:</strong> {error}
					</MessageBar>
				)}

				{/* Main content */}
				{targetUrl && !error && (
					<div style={{ position: 'relative', flex: 1, height: iframeStyle.height }}>
						<iframe
							ref={iframeRef}
							title={mode === 'bulk' ? 'Bulk Edit' : 'Edit Properties'}
							src={targetUrl}
							style={iframeStyle}
							onLoad={onIframeLoad}
							sandbox={sandbox}
							loading="lazy"
							onError={() => setError('Failed to load the edit form.')}
						/>

						{/* Interaction overlay */}
						{interactionState.isInteracting && (
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
									<Text variant="medium">{interactionState.message}</Text>
									{typeof interactionState.progress === 'number' && (
										<ProgressIndicator
											percentComplete={interactionState.progress / 100}
											styles={{ root: { width: '200px' } }}
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
						},
					}}
				>
					<Text variant="small">
						{mode === 'single'
							? 'Click Save to apply changes'
							: 'Use the details pane to edit in bulk'}
					</Text>
					<Text variant="small">Press Esc to close</Text>
				</Stack>
			</Stack>
		</Dialog>
	);
};
