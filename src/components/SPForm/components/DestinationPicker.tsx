import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
	Stack,
	Text,
	Dropdown,
	IDropdownOption,
	PrimaryButton,
	DefaultButton,
	Pivot,
	PivotItem,
	ComboBox,
	IComboBoxOption,
	MessageBar,
	MessageBarType,
	Label,
	SelectableOptionMenuItemType,
} from '@fluentui/react';
import { ContentTypeInfo, DestinationChoice, LibraryOption, PickerMode } from '../types';
import { SharePointService } from '../types';
import { normalizeError } from '../utils';

export interface DestinationPickerProps {
	pickerMode: PickerMode; // libraryFirst | contentTypeFirst | mixed

	libraries: LibraryOption[]; // configured library allow-list
	showContentTypePicker: boolean; // can hide CT UI entirely

	globalAllowedContentTypeIds?: string[] | 'all';

	spService: SharePointService; // for CT + title lookups

	// Prefills (optional)
	preselectContentTypeId?: string;

	onSubmit: (choice: DestinationChoice) => void;
	onCancel: () => void;

	// UI
	primaryText?: string;
	cancelText?: string;
	title?: string;
	subText?: string;
}

type TabKey = 'library' | 'contentType';
interface CTWithLib {
	ct: ContentTypeInfo;
	lib: LibraryOption;
}

export const DestinationPicker: React.FC<DestinationPickerProps> = (props) => {
	const {
		pickerMode,
		libraries,
		showContentTypePicker,
		globalAllowedContentTypeIds,
		spService,
		preselectContentTypeId,
		onSubmit,
		onCancel,
		primaryText = 'Continue',
		cancelText = 'Cancel',
		title = 'Choose destination',
		subText,
	} = props;

	const oneLibraryOnly = libraries.length === 1;

	const [activeTab, setActiveTab] = useState<TabKey>(() => {
		if (pickerMode === 'libraryFirst') return 'library';
		if (pickerMode === 'contentTypeFirst') return 'contentType';
		return preselectContentTypeId ? 'contentType' : 'library';
	});

	const [libTitles, setLibTitles] = useState<Record<string, string>>({});
	const [ctsByLib, setCtsByLib] = useState<Record<string, ContentTypeInfo[]>>({});
	const [loading, setLoading] = useState<boolean>(false);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	const [selectedLibraryUrl, setSelectedLibraryUrl] = useState<string>(
		oneLibraryOnly ? libraries[0].serverRelativeUrl : ''
	);
	const [selectedCTId, setSelectedCTId] = useState<string | undefined>(preselectContentTypeId);

	useEffect(() => {
		let disposed = false;
		const load = async () => {
			try {
				setLoading(true);
				setErrorMsg(null);
				const titles: Record<string, string> = { ...libTitles };

				for (const lib of libraries) {
					if (!titles[lib.serverRelativeUrl]) {
						try {
							titles[lib.serverRelativeUrl] =
								lib.label || (await spService.getLibraryTitle(lib.serverRelativeUrl));
						} catch {
							titles[lib.serverRelativeUrl] =
								lib.label || lib.serverRelativeUrl.split('/').pop() || lib.serverRelativeUrl;
						}
					}
				}

				const wantCTs = showContentTypePicker || pickerMode !== 'libraryFirst';
				const perLibCTs: Record<string, ContentTypeInfo[]> = { ...ctsByLib };

				if (wantCTs) {
					for (const lib of libraries) {
						if (!perLibCTs[lib.serverRelativeUrl]) {
							const raw = await spService.getLibraryContentTypes(lib.serverRelativeUrl);
							const visible = raw.filter((ct) => !ct.hidden);
							const filtered =
								lib.allowedContentTypeIds === 'all' || !lib.allowedContentTypeIds
									? visible
									: visible.filter((ct) => (lib.allowedContentTypeIds as string[]).includes(ct.id));
							const finalList =
								!globalAllowedContentTypeIds || globalAllowedContentTypeIds === 'all'
									? filtered
									: filtered.filter((ct) =>
											(globalAllowedContentTypeIds as string[]).includes(ct.id)
									  );
							perLibCTs[lib.serverRelativeUrl] = finalList.sort((a, b) =>
								a.name.localeCompare(b.name)
							);
						}
					}
				}

				if (disposed) return;
				setLibTitles(titles);
				setCtsByLib(perLibCTs);
			} catch (e) {
				if (disposed) return;
				setErrorMsg(normalizeError(e).message);
			} finally {
				if (!disposed) setLoading(false);
			}
		};

		load();
		return () => {
			disposed = true;
		};
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [libraries, showContentTypePicker, pickerMode, globalAllowedContentTypeIds]);

	// Build CT-first index
	const ctIndex: CTWithLib[] = useMemo(() => {
		if (!showContentTypePicker && pickerMode === 'libraryFirst') return [];
		const out: CTWithLib[] = [];
		for (const lib of libraries) {
			const cts = ctsByLib[lib.serverRelativeUrl] || [];
			for (const ct of cts) out.push({ ct, lib });
		}
		return out;
	}, [libraries, ctsByLib, pickerMode, showContentTypePicker]);

	// Library options
	const libraryOptions: IDropdownOption[] = useMemo(() => {
		return libraries.map((lib) => ({
			key: lib.serverRelativeUrl,
			text:
				libTitles[lib.serverRelativeUrl] ||
				lib.label ||
				lib.serverRelativeUrl.split('/').pop() ||
				lib.serverRelativeUrl,
		}));
	}, [libraries, libTitles]);

	// CT options for selected lib
	const ctOptionsForSelectedLibrary: IComboBoxOption[] = useMemo(() => {
		if (!selectedLibraryUrl) return [];
		const cts = ctsByLib[selectedLibraryUrl] || [];
		return cts.map((ct) => ({ key: ct.id, text: ct.name, data: ct }));
	}, [ctsByLib, selectedLibraryUrl]);

	// CT-first options (grouped by library header)
	const allCTGroupedOptions: IComboBoxOption[] = useMemo(() => {
		const grouped: IComboBoxOption[] = [];
		for (const lib of libraries) {
			const cts = ctsByLib[lib.serverRelativeUrl] || [];
			if (!cts.length) continue;
			const header: IComboBoxOption = {
				key: `hdr_${lib.serverRelativeUrl}`,
				text: libTitles[lib.serverRelativeUrl] || lib.label || lib.serverRelativeUrl,
				itemType: SelectableOptionMenuItemType.Header,
			} as any;
			grouped.push(header);
			for (const ct of cts) {
				grouped.push({
					key: ct.id,
					text: ct.name,
					data: { libUrl: lib.serverRelativeUrl, description: ct.description },
				});
			}
		}
		return grouped;
	}, [libraries, ctsByLib, libTitles]);

	// Validation
	const canContinue = useMemo(() => {
		if (activeTab === 'library') {
			if (!selectedLibraryUrl) return false;
			if (showContentTypePicker) {
				const cts = ctsByLib[selectedLibraryUrl] || [];
				if (cts.length === 1) return true;
				return !!selectedCTId;
			}
			return true;
		} else {
			return !!selectedCTId; // library auto-resolves from grouped option data
		}
	}, [activeTab, selectedLibraryUrl, selectedCTId, showContentTypePicker, ctsByLib]);

	// Auto-pick only CT for library-first
	useEffect(() => {
		if (activeTab !== 'library' || !selectedLibraryUrl || !showContentTypePicker) return;
		const cts = ctsByLib[selectedLibraryUrl] || [];
		if (cts.length === 1) setSelectedCTId(cts[0].id);
	}, [activeTab, selectedLibraryUrl, showContentTypePicker, ctsByLib]);

	const handleContinue = () => {
		let libUrl = selectedLibraryUrl;
		let ctId: string | undefined = selectedCTId;

		if (activeTab === 'contentType') {
			// When grouped, selected option stores libUrl in data
			const opt = allCTGroupedOptions.find((o) => o.key === selectedCTId);
			const data = opt?.data as any;
			if (data?.libUrl) libUrl = data.libUrl;
		} else if (activeTab === 'library' && showContentTypePicker) {
			const cts = ctsByLib[selectedLibraryUrl] || [];
			if (cts.length === 1) ctId = cts[0].id;
		}

		if (!libUrl) return;

		// compute friendly names
		const libraryTitle =
			libraries.find((l) => l.serverRelativeUrl === libUrl)?.label ||
			libTitles[libUrl] ||
			libUrl.split('/').pop() ||
			libUrl;

		const contentTypeName =
			(activeTab === 'library'
				? (ctsByLib[libUrl] || []).find((ct) => ct.id === ctId)?.name
				: // content-type-first (grouped)
				  (allCTGroupedOptions.find((o) => o.key === ctId)?.text as string)) || undefined;

		// include names in the payload (optional props)
		onSubmit({
			libraryUrl: libUrl,
			contentTypeId: ctId,
			libraryTitle,
			contentTypeName,
		});
	};

	// Render helpers
	const renderLibraryFirst = () => {
		const showLibraryDropdown = !oneLibraryOnly;
		return (
			<Stack tokens={{ childrenGap: 12 }}>
				{showLibraryDropdown ? (
					<>
						<Label>Document library</Label>
						<Dropdown
							options={libraryOptions}
							selectedKey={selectedLibraryUrl || undefined}
							onChange={(_, opt) => setSelectedLibraryUrl((opt?.key as string) || '')}
							placeholder="Select a library"
						/>
					</>
				) : (
					<Stack>
						<Label>Document library</Label>
						<Text>{libraryOptions[0]?.text}</Text>
					</Stack>
				)}
				{showContentTypePicker && (
					<>
						<Label>Content type</Label>
						{selectedLibraryUrl ? (
							<ComboBox
								allowFreeform={false}
								autoComplete="on"
								useComboBoxAsMenuWidth
								placeholder={
									(ctsByLib[selectedLibraryUrl]?.length ?? 0) > 0
										? 'Select a content type'
										: 'No content types available'
								}
								options={ctOptionsForSelectedLibrary}
								selectedKey={selectedCTId}
								onChange={(_, opt) => {
									// ignore group headers
									// @fluentui types: opt?.itemType === SelectableOptionMenuItemType.Header
									if ((opt as any)?.itemType === 1 /* Header */) return;
									setSelectedCTId(opt?.key as string);
								}}
							/>
						) : (
							<Text variant="small">Select a library first</Text>
						)}
						{selectedCTId && (
							<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
								{ctsByLib[selectedLibraryUrl]?.find((ct) => ct.id === selectedCTId)?.description ||
									''}
							</Text>
						)}
					</>
				)}
			</Stack>
		);
	};

	const renderContentTypeFirst = () => {
		const selectedOpt = allCTGroupedOptions.find((o) => o.key === selectedCTId);
		const selectedDesc = (selectedOpt?.data as any)?.description as string | undefined;

		return (
			<Stack tokens={{ childrenGap: 12 }}>
				<Label>Content type</Label>
				<ComboBox
					allowFreeform={false}
					autoComplete="on"
					useComboBoxAsMenuWidth
					placeholder={
						allCTGroupedOptions.length > 0 ? 'Select a content type' : 'No content types available'
					}
					options={allCTGroupedOptions}
					selectedKey={selectedCTId}
					onChange={(_, opt) => setSelectedCTId(opt?.key as string)}
				/>
				{selectedDesc && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						{selectedDesc}
					</Text>
				)}
			</Stack>
		);
	};

	const body =
		pickerMode === 'mixed' ? (
			<Pivot
				selectedKey={activeTab}
				onLinkClick={(item) => setActiveTab((item?.props.itemKey as TabKey) || 'library')}
			>
				<PivotItem headerText="By library" itemKey="library">
					<Stack styles={{ root: { marginTop: 12 } }}>{renderLibraryFirst()}</Stack>
				</PivotItem>
				<PivotItem headerText="By content type" itemKey="contentType">
					<Stack styles={{ root: { marginTop: 12 } }}>{renderContentTypeFirst()}</Stack>
				</PivotItem>
			</Pivot>
		) : activeTab === 'library' ? (
			renderLibraryFirst()
		) : (
			renderContentTypeFirst()
		);

	return (
		<Stack tokens={{ childrenGap: 16 }}>
			<Stack>
				<Text variant="xLarge">{title}</Text>
				{subText && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						{subText}
					</Text>
				)}
			</Stack>

			{errorMsg && (
				<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
					{errorMsg}
				</MessageBar>
			)}

			{body}

			<Stack
				horizontal
				horizontalAlign="end"
				tokens={{ childrenGap: 8 }}
				styles={{ root: { marginTop: 20 } }}
			>
				<DefaultButton text={cancelText} onClick={onCancel} />
				<PrimaryButton
					text={primaryText}
					onClick={handleContinue}
					disabled={loading || !canContinue}
				/>
			</Stack>
		</Stack>
	);
};
