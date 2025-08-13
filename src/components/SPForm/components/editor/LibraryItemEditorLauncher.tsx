// src/webparts/UploadAndEdit/components/editor/LibraryItemEditorLauncher.tsx
// SIMPLIFIED CLEAN VERSION - Core functionality only
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

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

const createAdvancedCSS = (hideBreadcrumbs: boolean, hideContentTypeField: boolean): string => {
	const css: string[] = [];

	css.push(`
		body {
			margin: 0 !important;
			padding: 0 !important;
			overflow-x: hidden !important;
		}
		.ms-Dialog-main { padding: 0 !important; }
	`);

	if (hideBreadcrumbs) {
		css.push(`
			.od-TopNav, .od-TopBar, .ms-CommandBar, .ms-Breadcrumb, .ms-Nav,
			[data-automationid="Breadcrumb"], [data-automationid="breadcrumb"],
			[class*="breadcrumb" i], [class*="topnav" i], [class*="navigation" i]
			{ display: none !important; }
		`);
	}

	if (hideContentTypeField) {
		css.push(`
			div[data-field="ContentType"], div[data-field="contenttype"],
			div[aria-label*="Content type" i], div[aria-label*="ContentType" i]
			{ display: none !important; }
		`);
	}

	return css.join('\n');
};

async function resolveListId(sp: SPFI, libraryServerRelativeUrl: string): Promise<string> {
	const list = await sp.web.getList(libraryServerRelativeUrl).select('Id')();
	return list?.Id as string;
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
				.select('Id', 'UniqueId', 'FileLeafRef', 'Modified', 'FileRef')();
			return {
				id: item.Id,
				uniqueId: item.UniqueId,
				fileName: item.FileLeafRef,
				modified: item.Modified,
				serverRelativeUrl: item.FileRef,
				listItemId: item.Id,
			};
		})
	);

	return results
		.filter((result): result is PromiseFulfilledResult<ItemInfo> => result.status === 'fulfilled')
		.map((result) => result.value);
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
		const itemIds = itemDetails.map((item) => item.id.toString());
		params.set('selectedItems', itemIds.join(','));
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
		hideBreadcrumbs = false,
		hideContentTypeField = false,
		sandboxExtra,
		autoHeightBestEffort = true,
	} = props;

	// ============================================================================
	// STATE
	// ============================================================================

	const [targetUrl, setTargetUrl] = useState<string>('');
	const [mode, setMode] = useState<'single' | 'bulk'>('single');
	const [modalOpen, setModalOpen] = useState<boolean>(isOpen && renderMode === 'modal');
	const [itemDetails, setItemDetails] = useState<ItemInfo[]>([]);
	const [error, setError] = useState<string | null>(null);
	const [isLoading, setIsLoading] = useState<boolean>(true);

	// ============================================================================
	// REFS AND MEMOIZED VALUES
	// ============================================================================

	const sp = useMemo(() => spfi(siteUrl).using(PnP_SPFX(spfxContext)), [siteUrl, spfxContext]);
	const iframeRef = useRef<HTMLIFrameElement | null>(null);
	const initRef = useRef<boolean>(false);

	// ============================================================================
	// INITIALIZATION
	// ============================================================================

	useEffect(() => {
		if (!itemIds?.length || initRef.current) return;

		initRef.current = true;

		const initialize = async () => {
			try {
				const single = itemIds.length === 1;
				setMode(single ? 'single' : 'bulk');

				if (single) {
					const listId = await resolveListId(sp, libraryServerRelativeUrl);
					const url = buildSingleEditUrl(siteUrl, listId, itemIds[0], window.location.href);
					setTargetUrl(url);

					if (onDetermined) {
						onDetermined({ mode: 'single', url, bulk: false });
					}
				} else {
					const details = await getItemDetails(sp, libraryServerRelativeUrl, itemIds);
					setItemDetails(details);

					const url = buildBulkViewUrl(siteUrl, libraryServerRelativeUrl, details, viewId);
					setTargetUrl(url);

					if (onDetermined) {
						onDetermined({ mode: 'bulk', url, bulk: true });
					}
				}

				// Handle non-modal modes
				if (renderMode === 'newtab') {
					window.open(targetUrl, '_blank');
					onDismiss?.();
					return;
				} else if (renderMode === 'samepage') {
					window.location.href = targetUrl;
					return;
				}

				setIsLoading(false);
			} catch (err) {
				console.error('Initialization failed:', err);
				setError(err instanceof Error ? err.message : 'Failed to initialize');
				setIsLoading(false);
			}
		};

		initialize();
	}, [itemIds, libraryServerRelativeUrl, siteUrl, viewId, renderMode, sp, onDetermined, onDismiss]);

	// ============================================================================
	// IFRAME HANDLING
	// ============================================================================

	const onIframeLoad = useCallback(() => {
		const frame = iframeRef.current;
		if (!frame || !targetUrl) return;

		console.log('Iframe loaded');

		if (onOpen) {
			onOpen({ mode, url: targetUrl });
		}

		// Inject CSS
		const doc = frame.contentDocument;
		if (doc && (hideBreadcrumbs || hideContentTypeField)) {
			try {
				const styleId = 'launcher-styles';
				if (!doc.getElementById(styleId)) {
					const style = doc.createElement('style');
					style.id = styleId;
					style.textContent = createAdvancedCSS(hideBreadcrumbs, hideContentTypeField);
					doc.head?.appendChild(style);
				}
			} catch (error) {
				console.warn('CSS injection failed:', error);
			}
		}

		// Handle single edit completion detection
		if (mode === 'single') {
			try {
				const href = frame.contentWindow?.location?.href;
				if (href && !isListFormEditUrl(href) && href.includes(window.location.origin)) {
					console.log('Edit completed, closing');
					setModalOpen(false);
					onSaved?.();
					onDismiss?.();
				}
			} catch (error) {
				// Cross-origin restrictions - ignore
			}
		}
	}, [mode, targetUrl, hideBreadcrumbs, hideContentTypeField, onOpen, onSaved, onDismiss]);

	// ============================================================================
	// MODAL VISIBILITY
	// ============================================================================

	useEffect(() => {
		if (renderMode === 'modal') {
			setModalOpen(!!isOpen);
		}
	}, [renderMode, isOpen]);

	// ============================================================================
	// RENDER
	// ============================================================================

	// Don't render for non-modal modes
	if (renderMode !== 'modal') return null;

	const dialogContentProps: IDialogContentProps = {
		type: DialogType.close,
		title:
			mode === 'bulk'
				? `Edit Properties - ${itemIds.length} file${itemIds.length > 1 ? 's' : ''}`
				: 'Edit File Properties',
		showCloseButton: true,
	};

	const iframeStyle: React.CSSProperties = {
		width: '100%',
		border: 'none',
		height: autoHeightBestEffort
			? `${Math.max(600, Math.floor(window.innerHeight * 0.85))}px`
			: '85vh',
		overflow: 'hidden',
	};

	const sandbox = `allow-scripts allow-same-origin allow-forms allow-popups allow-downloads allow-modals${
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
			}}
		>
			{error && (
				<MessageBar
					messageBarType={MessageBarType.error}
					onDismiss={() => setError(null)}
					actions={
						<IconButton
							iconProps={{ iconName: 'Refresh' }}
							title="Retry"
							onClick={() => window.location.reload()}
						/>
					}
				>
					<strong>Error:</strong> {error}
				</MessageBar>
			)}

			{isLoading && !error && (
				<Stack
					horizontalAlign="center"
					verticalAlign="center"
					tokens={{ childrenGap: 12 }}
					styles={{ root: { padding: '40px', minHeight: '200px' } }}
				>
					<Spinner size={SpinnerSize.large} />
					<Text>Loading editor...</Text>
				</Stack>
			)}

			{targetUrl && !error && !isLoading && (
				<iframe
					ref={iframeRef}
					title={mode === 'bulk' ? 'Bulk Edit Properties' : 'Edit File Properties'}
					src={targetUrl}
					style={iframeStyle}
					onLoad={onIframeLoad}
					onError={() => setError('Failed to load the edit form')}
					sandbox={sandbox}
				/>
			)}
		</Dialog>
	);
};
