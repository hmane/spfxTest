import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// ✅ Modern property pane (avoids deprecated APIs)
import {
	IPropertyPaneConfiguration,
	PropertyPaneDropdown,
	PropertyPaneToggle,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane';

// PnP property controls
import {
	PropertyFieldListPicker,
	PropertyFieldListPickerOrderBy,
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import {
	PropertyFieldCollectionData,
	CustomCollectionFieldType,
	IPropertyFieldCollectionDataProps,
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import { ToastHost } from './components/ToastHost';
import { UploadAndEditApp } from './components/UploadAndEditApp';
import { LibraryOption, UploadAndEditWebPartProps } from './types';

export interface IUploadAndEditWebPartProps extends UploadAndEditWebPartProps {
	librariesPicker?: any[];
	librariesExtras?: Array<{
		serverRelativeUrl: string;
		label?: string;
		defaultFolder?: string;
		minimalViewId?: string;
		allowedContentTypeIds?: string; // 'all' or comma-separated
	}>;
}

export default class UploadAndEditWebPart extends BaseClientSideWebPart<IUploadAndEditWebPartProps> {
	public render(): void {
		const libraries = this._composeLibraries();
		const configured = libraries.length > 0;

		const element: React.ReactElement = configured
			? React.createElement(
					ToastHost,
					null,
					React.createElement(UploadAndEditApp, {
						siteUrl: this.context.pageContext.web.absoluteUrl,
						spfxContext: this.context,

						pickerMode: this.properties.pickerMode ?? 'mixed',
						renderMode: this.properties.renderMode ?? 'modal',
						selectionScope: this.properties.selectionScope ?? 'multiple',
						showContentTypePicker: this.properties.showContentTypePicker ?? true,

						libraries,
						globalAllowedContentTypeIds: this.properties.globalAllowedContentTypeIds,

						overwritePolicy: this.properties.overwritePolicy ?? 'suffix',

						enableBulkAutoRefresh: this.properties.enableBulkAutoRefresh ?? true,
						bulkWatchAllItems: this.properties.bulkWatchAllItems ?? true,

						buttonLabel:
							this.properties.buttonLabel ??
							(this.properties.selectionScope === 'single' ? 'Upload file' : 'Upload files'),
						dropzoneHint:
							this.properties.dropzoneHint ??
							(this.properties.selectionScope === 'single'
								? 'Drop a file here'
								: 'Drop files here'),
						successToast: this.properties.successToast,

						disableDomNudges: this.properties.disableDomNudges ?? false,
						sandboxExtra: this.properties.sandboxExtra,

						showLoading: undefined,
						hideLoading: undefined,
						confirmOverwrite: undefined,
					})
			  )
			: React.createElement(
					'div',
					{ style: { padding: 12 } },
					'Configure this web part in the property pane.'
			  );

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

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
											title: 'Allowed CT IDs ("all" or comma-separated)',
											type: CustomCollectionFieldType.string,
										},
									],
								} as IPropertyFieldCollectionDataProps),
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
									value:
										this.properties.buttonLabel ??
										(this.properties.selectionScope === 'single' ? 'Upload file' : 'Upload files'),
								}),
								PropertyPaneTextField('dropzoneHint', {
									label: 'Dropzone hint',
									value:
										this.properties.dropzoneHint ??
										(this.properties.selectionScope === 'single'
											? 'Drop a file here'
											: 'Drop files here'),
								}),
								PropertyPaneTextField('successToast', {
									label: 'Success toast after saving properties',
									value: this.properties.successToast,
								}),
							],
						},
					],
				},
			],
		};
	}

	private async resolvePickedLibraries(): Promise<any[]> {
		if (!this.properties.librariesPicker || this.properties.librariesPicker.length === 0) {
			return [];
		}

		const sp = spfi().using(SPFx(this.context));

		// librariesPicker is array of GUIDs as strings
		const resolved = await Promise.all(
			(this.properties.librariesPicker as string[]).map(async (listId) => {
				try {
					const list = await sp.web.lists.getById(listId)();
					return {
						id: list.Id, // GUID
						title: list.Title,
						serverRelativeUrl: list.RootFolder.ServerRelativeUrl,
					};
				} catch (err) {
					console.error(`Error fetching library info for ${listId}`, err);
					return null;
				}
			})
		);

		return resolved.filter(Boolean);
	}

	/** Merge list picker selection with extras => LibraryOption[] */
	private _composeLibraries(): LibraryOption[] {
		const picked = (this.properties.librariesPicker || []) as Array<any>;
		const extras = (this.properties.librariesExtras || []).reduce((acc, r) => {
			acc[r.serverRelativeUrl] = r;
			return acc;
		}, {} as Record<string, any>);

		const libs: LibraryOption[] = picked.map((p: any) => {
			const ex = extras[p?.Url] || extras[p?.ServerRelativeUrl] || extras[p?.Title] || undefined;
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

		// dedupe
		const byUrl = new Map<string, LibraryOption>();
		for (const l of libs) {
			if (!l.serverRelativeUrl) continue;
			if (!byUrl.has(l.serverRelativeUrl)) byUrl.set(l.serverRelativeUrl, l);
			else {
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
