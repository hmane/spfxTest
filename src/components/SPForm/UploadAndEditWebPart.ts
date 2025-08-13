import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// Modern property pane (avoid deprecated symbols)
import {
	IPropertyPaneConfiguration,
	PropertyPaneDropdown,
	PropertyPaneToggle,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane';

// PnP SPFx property controls
import {
	PropertyFieldListPicker,
	PropertyFieldListPickerOrderBy,
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import {
	PropertyFieldCollectionData,
	CustomCollectionFieldType,
	IPropertyFieldCollectionDataProps,
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

// PnPjs to resolve list GUIDs → serverRelativeUrl
import { spfi, SPFI } from '@pnp/sp';
import { SPFx as PnP_SPFX } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders'; // for RootFolder expansion

import { ToastHost } from './components/ToastHost';
import { UploadAndEditApp } from './components/UploadAndEditApp';
import { LibraryOption, UploadAndEditWebPartProps } from './types';

export interface IUploadAndEditWebPartProps extends UploadAndEditWebPartProps {
	/** List GUIDs from ListPicker (multi-select) */
	librariesPicker?: string[];
	/** Per-library extra config entered in the CollectionData grid */
	librariesExtras?: Array<{
		serverRelativeUrl: string;
		label?: string;
		defaultFolder?: string;
		minimalViewId?: string;
		allowedContentTypeIds?: string; // 'all' or comma-separated
	}>;
}

export default class UploadAndEditWebPart extends BaseClientSideWebPart<IUploadAndEditWebPartProps> {
	private _sp!: SPFI;
	private _resolvedLibraries: LibraryOption[] = [];
	private _resolving = false;

	public async onInit(): Promise<void> {
		await super.onInit();
		this._sp = spfi().using(PnP_SPFX(this.context));
		await this._refreshLibraries(); // resolve on load
	}

	public render(): void {
		const configured = this._resolvedLibraries.length > 0;

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

						libraries: this._resolvedLibraries,
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

						// hook these if you have global loader / confirm dialog utils
						showLoading: undefined,
						hideLoading: undefined,
						confirmOverwrite: undefined,
					})
			  )
			: React.createElement(
					'div',
					{ style: { padding: 12, opacity: this._resolving ? 0.6 : 1 } },
					this._resolving ? 'Resolving libraries…' : 'Configure this web part in the property pane.'
			  );

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	// ---------------- Property Pane ----------------

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
									baseTemplate: 101, // Document Library
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

	/**
	 * When property pane changes, re-resolve libraries if needed.
	 */
	protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
		super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

		if (propertyPath === 'librariesPicker' || propertyPath === 'librariesExtras') {
			// Kick off re-resolution; re-render when done
			this._refreshLibraries().then(() => this.render());
		}
	}

	// ---------------- Resolution helpers ----------------

	/**
	 * Resolve GUIDs from ListPicker to LibraryOption[] with serverRelativeUrl,
	 * then merge with per-library extras.
	 */
	private async _refreshLibraries(): Promise<void> {
		this._resolving = true;
		try {
			const ids: string[] = Array.isArray(this.properties.librariesPicker)
				? (this.properties.librariesPicker as string[])
				: this.properties.librariesPicker
				? [this.properties.librariesPicker as unknown as string]
				: [];

			if (!ids.length) {
				this._resolvedLibraries = [];
				return;
			}

			// Resolve each list: Id, Title, RootFolder.ServerRelativeUrl
			const resolved = await Promise.all(
				ids.map(async (id) => {
					try {
						const list: any = await this._sp.web.lists
							.getById(id)
							.select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
							.expand('RootFolder')();

						return {
							id: list.Id as string,
							title: list.Title as string,
							serverRelativeUrl: list?.RootFolder?.ServerRelativeUrl as string,
						};
					} catch (e) {
						// swallow this one but log for admins
						// eslint-disable-next-line no-console
						console.warn('Failed to resolve library', id, e);
						return null;
					}
				})
			);

			const libsResolved = resolved.filter(Boolean) as Array<{
				id: string;
				title: string;
				serverRelativeUrl: string;
			}>;

			// Merge with extras
			const extrasArr = this.properties.librariesExtras || [];
			const extrasMap = new Map<string, any>();
			for (const ex of extrasArr) extrasMap.set(ex.serverRelativeUrl, ex);

			this._resolvedLibraries = libsResolved.map((r) => {
				const ex = extrasMap.get(r.serverRelativeUrl);
				const allowed =
					!ex?.allowedContentTypeIds || ex.allowedContentTypeIds.trim().toLowerCase() === 'all'
						? 'all'
						: ex.allowedContentTypeIds
								.split(',')
								.map((s: string) => s.trim())
								.filter(Boolean);

				const opt: LibraryOption = {
					serverRelativeUrl: r.serverRelativeUrl,
					label: ex?.label || r.title,
					defaultFolder: ex?.defaultFolder,
					minimalViewId: ex?.minimalViewId,
					allowedContentTypeIds: allowed,
				};
				return opt;
			});
		} finally {
			this._resolving = false;
		}
	}
}
