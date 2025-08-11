import { Version } from '@microsoft/sp-core-library';
import {
	BaseClientSideWebPart,
	IPropertyPaneConfiguration,
	PropertyPaneDropdown,
	PropertyPaneTextField,
	PropertyPaneToggle,
} from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
	PropertyFieldListPicker,
	PropertyFieldListPickerOrderBy,
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import {
	CustomCollectionFieldType,
	IPropertyFieldCollectionDataProps,
	PropertyFieldCollectionData,
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import { UploadAndEditApp } from './components/UploadAndEditApp';
import { LibraryOption, UploadAndEditWebPartProps } from './types';
import { ToastHost } from './components/ToastHost';

/** Web part props shape (persisted) */
export interface IUploadAndEditWebPartProps extends UploadAndEditWebPartProps {
	// the raw list picker selection (array of list IDs/urls depending on mode)
	librariesPicker?: any[];
	// the per-library extras (from CollectionData)
	librariesExtras?: Array<{
		serverRelativeUrl: string;
		label?: string;
		defaultFolder?: string;
		minimalViewId?: string;
		allowedContentTypeIds?: string; // comma-separated CT IDs or 'all'
	}>;
}

export default class UploadAndEditWebPart extends BaseClientSideWebPart<IUploadAndEditWebPartProps> {
	public render(): void {
		const libraries: LibraryOption[] = this._composeLibraries();

		const element = React.createElement(
			ToastHost,
			null,
			React.createElement(UploadAndEditApp, {
				// context + site
				siteUrl: this.context.pageContext.web.absoluteUrl,
				spfxContext: this.context,

				// config
				pickerMode: this.properties.pickerMode ?? 'mixed',
				renderMode: this.properties.renderMode ?? 'modal',
				selectionScope: this.properties.selectionScope ?? 'multiple',
				allowFolderSelection: this.properties.allowFolderSelection ?? true,
				showContentTypePicker: this.properties.showContentTypePicker ?? true,
				overwritePolicy: this.properties.overwritePolicy ?? 'suffix',

				libraries,
				defaultLibrary: this.properties.defaultLibrary,
				globalAllowedContentTypeIds: this.properties.globalAllowedContentTypeIds,

				enableBulkAutoRefresh: this.properties.enableBulkAutoRefresh ?? true,
				bulkWatchAllItems: this.properties.bulkWatchAllItems ?? true,

				buttonLabel: this.properties.buttonLabel ?? 'Upload files',
				dropzoneHint:
					this.properties.dropzoneHint ?? 'Drag & drop files here, or click Select files',
				successToast: this.properties.successToast,

				disableDomNudges: this.properties.disableDomNudges ?? false,
				sandboxExtra: this.properties.sandboxExtra,

				minimalViewId: undefined, // prefer per-library minimalViewId
			})
		);

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	// --------- Property pane ---------

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: { description: 'Upload + Edit Web Part' },
					groups: [
						{
							groupName: 'Mode & Behavior',
							groupFields: [
								PropertyPaneDropdown('pickerMode', {
									label: 'Picker mode',
									options: [
										{ key: 'libraryFirst', text: 'Library-first' },
										{ key: 'contentTypeFirst', text: 'Content type-first' },
										{ key: 'mixed', text: 'Mixed (tabs)' },
									],
									selectedKey: this.properties.pickerMode ?? 'mixed',
								}),
								PropertyPaneDropdown('renderMode', {
									label: 'Edit form render target',
									options: [
										{ key: 'modal', text: 'Modal (recommended)' },
										{ key: 'samepage', text: 'Same page' },
										{ key: 'newtab', text: 'New tab' },
									],
									selectedKey: this.properties.renderMode ?? 'modal',
								}),
								PropertyPaneDropdown('selectionScope', {
									label: 'Allow multiple files per batch?',
									options: [
										{ key: 'single', text: 'Single file' },
										{ key: 'multiple', text: 'Multiple files' },
									],
									selectedKey: this.properties.selectionScope ?? 'multiple',
								}),
								PropertyPaneToggle('allowFolderSelection', {
									label: 'Allow choosing a folder',
									onText: 'Yes',
									offText: 'No',
									checked: this.properties.allowFolderSelection ?? true,
								}),
								PropertyPaneToggle('showContentTypePicker', {
									label: 'Show content type picker',
									onText: 'Yes',
									offText: 'No',
									checked: this.properties.showContentTypePicker ?? true,
								}),
								PropertyPaneDropdown('overwritePolicy', {
									label: 'On filename conflict',
									options: [
										{ key: 'overwrite', text: 'Overwrite' },
										{ key: 'skip', text: 'Skip and report' },
										{ key: 'suffix', text: 'Add numeric suffix (recommended)' },
									],
									selectedKey: this.properties.overwritePolicy ?? 'suffix',
								}),
							],
						},
						{
							groupName: 'Allowed libraries',
							groupFields: [
								// 1) Pick the document libraries (BaseTemplate 101)
								PropertyFieldListPicker('librariesPicker', {
									label: 'Choose one or more document libraries',
									selectedList: (this.properties.librariesPicker as any) || [],
									includeHidden: false,
									baseTemplate: 101,
									orderBy: PropertyFieldListPickerOrderBy.Title,
									multiSelect: true,
									onPropertyChange: this.onPropertyPaneFieldChanged,
									properties: this.properties,
									context: this.context as any,
									deferredValidationTime: 200,
									key: 'librariesPicker',
								}) as any,

								// 2) Optional: per-library metadata / overrides
								PropertyFieldCollectionData('librariesExtras', {
									key: 'librariesExtras',
									label: 'Per-library options (optional)',
									panelHeader: 'Library options',
									manageBtnLabel: 'Edit libraries options',
									value: this.properties.librariesExtras || [],
									fields: [
										{
											id: 'serverRelativeUrl',
											title: 'Library URL (server-relative)',
											type: CustomCollectionFieldType.string,
											required: true,
										},
										{
											id: 'label',
											title: 'Friendly label',
											type: CustomCollectionFieldType.string,
										},
										{
											id: 'defaultFolder',
											title: 'Default subfolder',
											type: CustomCollectionFieldType.string,
										},
										{
											id: 'minimalViewId',
											title: 'Minimal View ID (no braces)',
											type: CustomCollectionFieldType.string,
										},
										{
											id: 'allowedContentTypeIds',
											title: 'Allowed CT IDs (comma or "all")',
											type: CustomCollectionFieldType.string,
										},
									],
								} as IPropertyFieldCollectionDataProps),
								PropertyPaneTextField('defaultLibrary', {
									label: 'Default library (server-relative URL)',
									description: 'Optional. If only one library is configured, this is ignored.',
									value: this.properties.defaultLibrary,
								}),
							],
						},
						{
							groupName: 'Editor behavior',
							groupFields: [
								PropertyPaneToggle('enableBulkAutoRefresh', {
									label: 'Auto-close bulk editor when items change',
									onText: 'On',
									offText: 'Off',
									checked: this.properties.enableBulkAutoRefresh ?? true,
								}),
								PropertyPaneToggle('bulkWatchAllItems', {
									label: 'Watch all selected items (not just first)',
									onText: 'On',
									offText: 'Off',
									checked: this.properties.bulkWatchAllItems ?? true,
								}),
								PropertyPaneToggle('disableDomNudges', {
									label: 'Disable DOM nudges (pane/save)',
									onText: 'On',
									offText: 'Off',
									checked: this.properties.disableDomNudges ?? false,
								}),
								PropertyPaneTextField('sandboxExtra', {
									label: 'Extra iframe sandbox flags (optional)',
									description: 'Appends: allow-scripts allow-same-origin allow-forms allow-popups',
									value: this.properties.sandboxExtra,
								}),
							],
						},
						{
							groupName: 'UI text',
							groupFields: [
								PropertyPaneTextField('buttonLabel', {
									label: 'Upload button title',
									value: this.properties.buttonLabel ?? 'Upload files',
								}),
								PropertyPaneTextField('dropzoneHint', {
									label: 'Dropzone hint',
									value:
										this.properties.dropzoneHint ?? 'Drag & drop files here, or click Select files',
								}),
								PropertyPaneTextField('successToast', {
									label: 'Success message (optional)',
									value: this.properties.successToast,
								}),
							],
						},
					],
				},
			],
		};
	}

	// ---------- Helpers ----------

	/**
	 * Merge the list-picker selection with the per-library extras
	 * to produce the final LibraryOption[] used by the app.
	 */
	private _composeLibraries(): LibraryOption[] {
		const picked = (this.properties.librariesPicker || []) as Array<any>;
		const extras = (this.properties.librariesExtras || []).reduce((acc, r) => {
			acc[r.serverRelativeUrl] = r;
			return acc;
		}, {} as Record<string, any>);

		// PropertyFieldListPicker returns items with .Id and .Title and .Url (web-relative)
		// We want server-relative library URLs. In SPFx, list root folder ServerRelativeUrl
		// isn’t provided by the picker directly, so we’ll accept either server-relative paths
		// from `librariesExtras` OR fall back to the picker’s `Url` (often web-relative).
		// In runtime, DestinationPicker resolves titles again from SharePoint.
		const libs: LibraryOption[] = picked.map((p: any) => {
			// Try extras first (authoritative for URL)
			const ex = extras[p?.Url] || extras[p?.ServerRelativeUrl] || extras[p?.Title] || undefined;

			// Best-effort server-relative: if picker gives '/sites/...', use that; else assume it’s server-relative already
			const serverRelativeUrl =
				ex?.serverRelativeUrl ||
				(p?.Url?.startsWith('/') ? p.Url : p?.ServerRelativeUrl) ||
				p?.Url ||
				'';

			const allowedContentTypeIds =
				!ex?.allowedContentTypeIds || ex.allowedContentTypeIds.trim().toLowerCase() === 'all'
					? 'all'
					: ex.allowedContentTypeIds
							.split(',')
							.map((s: string) => s.trim())
							.filter(Boolean);

			return {
				serverRelativeUrl,
				label: ex?.label || p?.Title,
				defaultFolder: ex?.defaultFolder,
				minimalViewId: ex?.minimalViewId,
				allowedContentTypeIds,
			} as LibraryOption;
		});

		// Deduplicate by serverRelativeUrl
		const byUrl = new Map<string, LibraryOption>();
		for (const l of libs) {
			if (!l.serverRelativeUrl) continue;
			if (!byUrl.has(l.serverRelativeUrl)) byUrl.set(l.serverRelativeUrl, l);
			else {
				// merge extras if needed (first pick wins; add missing props)
				const cur = byUrl.get(l.serverRelativeUrl)!;
				byUrl.set(l.serverRelativeUrl, {
					...cur,
					label: cur.label || l.label,
					defaultFolder: cur.defaultFolder || l.defaultFolder,
					minimalViewId: cur.minimalViewId || l.minimalViewId,
					allowedContentTypeIds: cur.allowedContentTypeIds ?? l.allowedContentTypeIds,
				});
			}
		}
		return Array.from(byUrl.values());
	}
}
