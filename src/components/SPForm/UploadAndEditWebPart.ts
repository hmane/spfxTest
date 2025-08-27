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
	PropertyPaneHorizontalRule,
	PropertyPaneLabel,
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

	private _normalizeGlobalCTs(): string[] | 'all' | undefined {
		const raw = this.properties.globalAllowedContentTypeIds as any;
		if (!raw) return undefined;
		if (Array.isArray(raw)) return raw;
		const s = String(raw).trim();
		if (!s) return undefined;
		if (s.toLowerCase() === 'all') return 'all';
		return s
			.split(',')
			.map((p) => p.trim())
			.filter(Boolean);
	}

	public render(): void {
		const configured = this._resolvedLibraries.length > 0;
		const globalCTs = this._normalizeGlobalCTs();

		// Always render ToastHost as the root component to maintain consistent DOM structure
		const element: React.ReactElement = React.createElement(
			ToastHost,
			null,
			configured
				? React.createElement(UploadAndEditApp, {
						siteUrl: this.context.pageContext.web.absoluteUrl,
						spfxContext: this.context,

						pickerMode: this.properties.pickerMode ?? 'mixed',
						renderMode: this.properties.renderMode ?? 'modal',
						selectionScope: this.properties.selectionScope ?? 'multiple',
						showContentTypePicker: this.properties.showContentTypePicker ?? true,

						libraries: this._resolvedLibraries,
						globalAllowedContentTypeIds: globalCTs,

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
				: React.createElement(
						'div',
						{
							key: 'configuration-placeholder', // Add key for better React reconciliation
							style: {
								padding: 20,
								textAlign: 'center',
								opacity: this._resolving ? 0.6 : 1,
								border: '1px dashed #ccc',
								borderRadius: 4,
								backgroundColor: '#f9f9f9',
							},
						},
						React.createElement(
							'div',
							{ style: { fontSize: 16, marginBottom: 8 } },
							this._resolving ? '⏳ Resolving libraries…' : '⚙️ Configuration Required'
						),
						React.createElement(
							'div',
							{ style: { fontSize: 14, color: '#666' } },
							this._resolving
								? 'Please wait while we load your libraries.'
								: 'Configure this web part in the property pane to get started.'
						)
					)
		);

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.1');
	}

	// ---------------- Enhanced Property Pane ----------------

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description:
							'Configure the core behavior and appearance of your Upload + Edit web part.',
					},
					groups: [
						{
							groupName: 'Core Behavior',
							groupFields: [
								PropertyPaneDropdown('pickerMode', {
									label: 'Picker mode',
									options: [
										{
											key: 'libraryFirst',
											text: '📁 Library-first - Select library, then content type',
										},
										{
											key: 'contentTypeFirst',
											text: '📋 Content type-first - Select content type, then library',
										},
										{ key: 'mixed', text: '🔀 Mixed (tabs) - Show both options in tabs' },
									],
									selectedKey: this.properties.pickerMode ?? 'mixed',
								}),
								PropertyPaneDropdown('renderMode', {
									label: 'Edit form display',
									options: [
										{ key: 'modal', text: '🪟 Modal dialog (recommended)' },
										{ key: 'samepage', text: '📄 Same page' },
										{ key: 'newtab', text: '🔗 New browser tab' },
									],
									selectedKey: this.properties.renderMode ?? 'modal',
								}),
								PropertyPaneDropdown('selectionScope', {
									label: 'File selection scope',
									options: [
										{ key: 'single', text: 'Single file only' },
										{ key: 'multiple', text: 'Multiple files allowed' },
									],
									selectedKey: this.properties.selectionScope ?? 'multiple',
								}),
								PropertyPaneToggle('showContentTypePicker', {
									label: 'Show content type picker',
									onText: 'Enabled',
									offText: 'Disabled',
									checked: this.properties.showContentTypePicker ?? true,
								}),
								PropertyPaneDropdown('overwritePolicy', {
									label: 'File conflict resolution',
									options: [
										{ key: 'overwrite', text: '🔄 Overwrite existing files' },
										{ key: 'skip', text: '⏭️ Skip duplicate files and report' },
										{ key: 'suffix', text: '🔢 Add numeric suffix (recommended)' },
									],
									selectedKey: this.properties.overwritePolicy ?? 'suffix',
								}),
							],
						},
						{
							groupName: 'Document Libraries',
							groupFields: [
								PropertyFieldListPicker('librariesPicker', {
									label: 'Select document libraries',
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
									disabled: false,
								}) as any,
								PropertyPaneLabel('librariesHelp', {
									text: 'Choose one or more document libraries where users can upload files. Additional per-library settings can be configured below.',
								}),
								PropertyFieldCollectionData('librariesExtras', {
									key: 'librariesExtras',
									label: 'Per-library configuration (optional)',
									panelHeader: 'Library-specific settings',
									manageBtnLabel: 'Configure library options',
									value: this.properties.librariesExtras || [],
									fields: [
										{
											id: 'serverRelativeUrl',
											title: 'Library URL (server-relative)',
											type: CustomCollectionFieldType.string,
											required: true,
											placeholder: '/sites/mysite/LibraryName',
										},
										{
											id: 'label',
											title: 'Display name (optional)',
											type: CustomCollectionFieldType.string,
											placeholder: 'My Documents',
										},
										{
											id: 'defaultFolder',
											title: 'Default subfolder (optional)',
											type: CustomCollectionFieldType.string,
											placeholder: 'Subfolder/Path',
										},
										{
											id: 'minimalViewId',
											title: 'Minimal View ID (optional, no braces)',
											type: CustomCollectionFieldType.string,
											placeholder: '12345678-1234-1234-1234-123456789012',
										},
										{
											id: 'allowedContentTypeIds',
											title: 'Allowed Content Type IDs (optional)',
											type: CustomCollectionFieldType.string,
											placeholder: 'all or 0x0101001234567890...,0x0101001234567891...',
										},
									],
									panelDescription:
										'Configure specific settings for each library. Only the server-relative URL is required - other settings are optional and will override global defaults.',
								} as IPropertyFieldCollectionDataProps),
							],
						},
						{
							groupName: 'User Interface Text',
							groupFields: [
								PropertyPaneTextField('buttonLabel', {
									label: 'Upload button text',
									value:
										this.properties.buttonLabel ||
										(this.properties.selectionScope === 'single' ? 'Upload file' : 'Upload files'),
									placeholder:
										this.properties.selectionScope === 'single' ? 'Upload file' : 'Upload files',
								}),
								PropertyPaneTextField('dropzoneHint', {
									label: 'Drag & drop hint text',
									value:
										this.properties.dropzoneHint ||
										(this.properties.selectionScope === 'single'
											? 'Drop a file here'
											: 'Drop files here'),
									placeholder:
										this.properties.selectionScope === 'single'
											? 'Drop a file here'
											: 'Drop files here',
								}),
								PropertyPaneTextField('successToast', {
									label: 'Success notification message (optional)',
									value: this.properties.successToast || '',
									placeholder: 'Files uploaded successfully!',
								}),
								PropertyPaneLabel('textHelp', {
									text: 'Customize the text displayed to users. Button and dropzone text will update automatically based on single/multiple file selection.',
								}),
							],
						},
					],
				},
				// Second page for advanced settings
				{
					header: {
						description: 'Advanced configuration options for power users and administrators.',
					},
					groups: [
						{
							groupName: 'Editor Behavior',
							groupFields: [
								PropertyPaneToggle('enableBulkAutoRefresh', {
									label: 'Auto-close bulk editor when items change',
									onText: 'Enabled',
									offText: 'Disabled',
									checked: this.properties.enableBulkAutoRefresh ?? true,
								}),
								PropertyPaneToggle('bulkWatchAllItems', {
									label: 'Watch all selected items (not just first)',
									onText: 'Watch all items',
									offText: 'Watch first item only',
									checked: this.properties.bulkWatchAllItems ?? true,
									disabled: !this.properties.enableBulkAutoRefresh,
								}),
								PropertyPaneLabel('editorHelp', {
									text: 'Auto-refresh closes the bulk editor automatically when items are modified. Watching all items provides better user feedback but may impact performance with many files.',
								}),
							],
						},
						{
							groupName: 'Advanced Technical Settings',
							groupFields: [
								PropertyPaneToggle('disableDomNudges', {
									label: 'Disable DOM nudges (pane/save)',
									onText: 'Disabled',
									offText: 'Enabled',
									checked: this.properties.disableDomNudges ?? false,
								}),
								PropertyPaneTextField('sandboxExtra', {
									label: 'Extra iframe sandbox flags (optional)',
									description:
										'Advanced setting - additional sandbox permissions for iframe-based editors',
									value: this.properties.sandboxExtra || '',
									placeholder: 'allow-downloads allow-modals',
									multiline: true,
									rows: 2,
								}),
								PropertyPaneTextField('globalAllowedContentTypeIds', {
									label: 'Global allowed content type IDs (optional)',
									description:
										'Comma-separated list of content type IDs to restrict globally, or leave empty for no restrictions',
									value:
										typeof this.properties.globalAllowedContentTypeIds === 'string'
											? this.properties.globalAllowedContentTypeIds
											: Array.isArray(this.properties.globalAllowedContentTypeIds)
											? this.properties.globalAllowedContentTypeIds.join(',')
											: '',
									placeholder: '0x0101001234567890,0x0101001234567891',
									multiline: true,
									rows: 2,
								}),
								PropertyPaneHorizontalRule(),
								PropertyPaneLabel('technicalHelp', {
									text: '⚠️ These are advanced settings that should only be modified by experienced administrators. Incorrect values may cause the web part to malfunction.',
								}),
							],
						},
					],
				},
			],
		};
	}

	/**
	 * Enhanced property change handler with validation and conditional updates
	 */
	protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
		super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

		// Handle dependent property updates
		if (propertyPath === 'enableBulkAutoRefresh' && !newValue) {
			// If auto-refresh is disabled, also disable watching all items
			this.properties.bulkWatchAllItems = false;
		}

		if (propertyPath === 'selectionScope') {
			// Update button and dropzone text when selection scope changes
			if (newValue === 'single') {
				if (!this.properties.buttonLabel || this.properties.buttonLabel === 'Upload files') {
					this.properties.buttonLabel = 'Upload file';
				}
				if (!this.properties.dropzoneHint || this.properties.dropzoneHint === 'Drop files here') {
					this.properties.dropzoneHint = 'Drop a file here';
				}
			} else {
				if (!this.properties.buttonLabel || this.properties.buttonLabel === 'Upload file') {
					this.properties.buttonLabel = 'Upload files';
				}
				if (!this.properties.dropzoneHint || this.properties.dropzoneHint === 'Drop a file here') {
					this.properties.dropzoneHint = 'Drop files here';
				}
			}
		}

		// Trigger library resolution if needed
		if (propertyPath === 'librariesPicker' || propertyPath === 'librariesExtras') {
			this._refreshLibraries().then(() => this.render());
		}

		// Re-render for UI-affecting changes
		const uiProperties = ['buttonLabel', 'dropzoneHint', 'successToast'];
		if (uiProperties.includes(propertyPath)) {
			this.render();
		}

		// Refresh property pane to show/hide conditional fields
		if (propertyPath === 'enableBulkAutoRefresh' || propertyPath === 'selectionScope') {
			this.context.propertyPane.refresh();
		}
	}

	/**
	 * Validate configuration when property pane opens
	 */
	protected onPropertyPaneConfigurationStart(): void {
		this._refreshLibraries();
	}

	/**
	 * Show validation results when configuration is complete
	 */
	protected onPropertyPaneConfigurationComplete(): void {
		const validationResults = this._validateConfiguration();

		if (validationResults.length > 0) {
			console.warn('Web part configuration validation warnings:', validationResults);
		}

		// Show success message if configured
		if (this.properties.successToast && this._resolvedLibraries.length > 0) {
			console.log('Configuration saved successfully:', this.properties.successToast);
		}
	}

	// ---------------- Enhanced Resolution helpers ----------------

	/**
	 * Resolve GUIDs from ListPicker to LibraryOption[] with serverRelativeUrl,
	 * then merge with per-library extras with enhanced error handling.
	 */
	private async _refreshLibraries(): Promise<void> {
		this._resolving = true;
		this.render(); // Show loading state

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

			// Resolve each list with better error handling
			const resolved = await Promise.allSettled(
				ids.map(async (id) => {
					const list: any = await this._sp.web.lists
						.getById(id)
						.select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
						.expand('RootFolder')();

					return {
						id: list.Id as string,
						title: list.Title as string,
						serverRelativeUrl: list?.RootFolder?.ServerRelativeUrl as string,
					};
				})
			);

			// Filter successful results and log failures
			const libsResolved = resolved
				.filter((result, index) => {
					if (result.status === 'rejected') {
						console.warn(`Failed to resolve library ${ids[index]}:`, result.reason);
						return false;
					}
					return true;
				})
				.map((result: any) => result.value);

			// Merge with extras configuration
			const extrasArr = this.properties.librariesExtras || [];
			const extrasMap = new Map<string, any>();
			for (const ex of extrasArr) {
				extrasMap.set(ex.serverRelativeUrl, ex);
			}

			this._resolvedLibraries = libsResolved.map((r) => {
				const ex = extrasMap.get(r.serverRelativeUrl);

				// Handle allowedContentTypeIds with proper type checking
				let allowed: string[] | 'all' | undefined;
				if (ex?.allowedContentTypeIds) {
					const ctIds = ex.allowedContentTypeIds.trim();
					if (ctIds.toLowerCase() === 'all') {
						allowed = 'all';
					} else {
						allowed = ctIds
							.split(',')
							.map((s: string) => s.trim())
							.filter(Boolean);
					}
				} else {
					allowed = undefined;
				}

				const opt: LibraryOption = {
					serverRelativeUrl: r.serverRelativeUrl,
					label: ex?.label || r.title,
					defaultFolder: ex?.defaultFolder,
					minimalViewId: ex?.minimalViewId,
					allowedContentTypeIds: allowed,
				};
				return opt;
			});
		} catch (error) {
			console.error('Error refreshing libraries:', error);
			this._resolvedLibraries = [];
		} finally {
			this._resolving = false;
			this.render();
		}
	}

	/**
	 * Validate web part configuration and return any issues
	 */
	private _validateConfiguration(): string[] {
		const issues: string[] = [];

		// Check if libraries are configured
		if (!this._resolvedLibraries.length) {
			issues.push('At least one document library must be selected');
		}

		// Validate globalAllowedContentTypeIds format if provided
		if (this.properties.globalAllowedContentTypeIds) {
			// Handle both string and string array types
			const globalCtIds =
				typeof this.properties.globalAllowedContentTypeIds === 'string'
					? this.properties.globalAllowedContentTypeIds.split(',').map((id) => id.trim())
					: this.properties.globalAllowedContentTypeIds;

			const invalidIds = globalCtIds.filter((id) => id && !id.match(/^0x[0-9A-Fa-f]+$/));
			if (invalidIds.length > 0) {
				issues.push(`Invalid content type ID format: ${invalidIds.join(', ')}`);
			}
		}

		// Validate library extras configuration
		if (this.properties.librariesExtras) {
			this.properties.librariesExtras.forEach((extra, index) => {
				if (!extra.serverRelativeUrl) {
					issues.push(`Library configuration ${index + 1}: Server-relative URL is required`);
				} else if (!extra.serverRelativeUrl.startsWith('/')) {
					issues.push(
						`Library configuration ${index + 1}: Server-relative URL must start with '/'`
					);
				}

				if (extra.minimalViewId && !extra.minimalViewId.match(/^[0-9A-Fa-f-]{36}$/)) {
					issues.push(`Library configuration ${index + 1}: Invalid view ID format`);
				}

				// Validate allowedContentTypeIds in extras
				if (extra.allowedContentTypeIds && extra.allowedContentTypeIds.toLowerCase() !== 'all') {
					const ctIds = extra.allowedContentTypeIds.split(',').map((id) => id.trim());
					const invalidIds = ctIds.filter((id) => id && !id.match(/^0x[0-9A-Fa-f]+$/));
					if (invalidIds.length > 0) {
						issues.push(
							`Library configuration ${
								index + 1
							}: Invalid content type ID format: ${invalidIds.join(', ')}`
						);
					}
				}
			});
		}

		return issues;
	}

	/**
	 * Get default properties for new web part instances
	 */
	protected getDefaultProperties(): Partial<IUploadAndEditWebPartProps> {
		return {
			pickerMode: 'mixed',
			renderMode: 'modal',
			selectionScope: 'multiple',
			showContentTypePicker: true,
			overwritePolicy: 'suffix',
			enableBulkAutoRefresh: true,
			bulkWatchAllItems: true,
			disableDomNudges: false,
			buttonLabel: 'Upload files',
			dropzoneHint: 'Drop files here',
		};
	}
}
