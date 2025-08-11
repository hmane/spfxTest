import {
    ComboBox,
    DefaultButton,
    Dropdown,
    IComboBoxOption,
    IDropdownOption,
    Label,
    MessageBar,
    MessageBarType,
    Pivot,
    PivotItem,
    PrimaryButton,
    Stack,
    Text,
    TextField,
} from '@fluentui/react';
import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
    ContentTypeInfo,
    DestinationChoice,
    LibraryOption,
    PickerMode,
    SharePointService,
    UploadSelectionScope
} from '../types';
import { isNonEmptyString, normalizeError, trimLeadingSlash } from '../utils';

export interface DestinationPickerProps {
	pickerMode: PickerMode; // libraryFirst | contentTypeFirst | mixed
	selectionScope: UploadSelectionScope; // single | multiple (future: could restrict CTs if needed)

	libraries: LibraryOption[]; // configured library allow-list
	defaultLibrary?: string; // optional default lib (server-relative)
	showContentTypePicker: boolean; // show/hide CT UI
	allowFolderSelection: boolean; // show/hide folder field

	globalAllowedContentTypeIds?: string[] | 'all';

	spService: SharePointService; // for CT + title lookups

	// Prefills (optional)
	preselectContentTypeId?: string;

	onSubmit: (choice: DestinationChoice) => void;
	onCancel: () => void;

	// UI
	primaryText?: string; // e.g., "Continue"
	cancelText?: string; // e.g., "Cancel"
	title?: string; // e.g., "Select destination"
	subText?: string; // optional helper text
}

type TabKey = 'library' | 'contentType';

interface CTWithLib {
	ct: ContentTypeInfo;
	lib: LibraryOption;
}

export const DestinationPicker: React.FC<DestinationPickerProps> = (props) => {
	const {
		pickerMode,
		selectionScope,

		libraries,
		defaultLibrary,
		showContentTypePicker,
		allowFolderSelection,

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

	// ---------- Derived basics ----------
	const oneLibraryOnly = libraries.length === 1;
	const initialLibUrl = oneLibraryOnly ? libraries[0].serverRelativeUrl : defaultLibrary || '';

	const [activeTab, setActiveTab] = useState<TabKey>(() => {
		if (pickerMode === 'libraryFirst') return 'library';
		if (pickerMode === 'contentTypeFirst') return 'contentType';
		// mixed: choose based on what’s preselected
		return preselectContentTypeId ? 'contentType' : 'library';
	});

	const [libTitles, setLibTitles] = useState<Record<string, string>>({});
	const [ctsByLib, setCtsByLib] = useState<Record<string, ContentTypeInfo[]>>({});
	const [loading, setLoading] = useState<boolean>(false);
	const [errorMsg, setErrorMsg] = useState<string | null>(null);

	// Selections
	const [selectedLibraryUrl, setSelectedLibraryUrl] = useState<string>(initialLibUrl);
	const [selectedCTId, setSelectedCTId] = useState<string | undefined>(preselectContentTypeId);
	const [folderPath, setFolderPath] = useState<string>('');

	// ---------- Load library titles + CTs lazily ----------
	useEffect(() => {
		let disposed = false;

		const load = async () => {
			try {
				setLoading(true);
				setErrorMsg(null);

				// Titles (only for those we’ll show)
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

				// If we need CTs at all:
				const wantCTs = showContentTypePicker || pickerMode !== 'libraryFirst';
				const perLibCTs: Record<string, ContentTypeInfo[]> = { ...ctsByLib };

				if (wantCTs) {
					for (const lib of libraries) {
						if (!perLibCTs[lib.serverRelativeUrl]) {
							const raw = await spService.getLibraryContentTypes(lib.serverRelativeUrl);
							const visible = raw.filter((ct) => !ct.hidden); // already filtered in service, but extra guard
							// Respect per-library allowed CTs if configured
							const filtered =
								lib.allowedContentTypeIds === 'all' || !lib.allowedContentTypeIds
									? visible
									: visible.filter((ct) => lib.allowedContentTypeIds!.includes(ct.id));

							// Respect global allow-list if provided
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

	// ---------- Build CT-first search index ----------
	const ctIndex: CTWithLib[] = useMemo(() => {
		if (!showContentTypePicker && pickerMode === 'libraryFirst') return [];
		const out: CTWithLib[] = [];
		for (const lib of libraries) {
			const cts = ctsByLib[lib.serverRelativeUrl] || [];
			for (const ct of cts) {
				out.push({ ct, lib });
			}
		}
		// Deduplicate (CT might appear in multiple libs with same id & name; we keep multiple entries but annotate)
		return out;
	}, [libraries, ctsByLib, pickerMode, showContentTypePicker]);

	// ---------- Library options ----------
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

	// ---------- CT options for the chosen library ----------
	const ctOptionsForSelectedLibrary: IComboBoxOption[] = useMemo(() => {
		if (!selectedLibraryUrl) return [];
		const cts = ctsByLib[selectedLibraryUrl] || [];
		return cts.map((ct) => ({
			key: ct.id,
			text: ct.name,
			data: ct,
		}));
	}, [ctsByLib, selectedLibraryUrl]);

	// ---------- CT-first options (cross-library) ----------
	const allCTOptions: IComboBoxOption[] = useMemo(() => {
		// Show CT name; if it appears in >1 libraries, append a chip-like suffix
		const counts: Record<string, number> = {};
		for (const e of ctIndex) counts[e.ct.id] = (counts[e.ct.id] || 0) + 1;

		// Using first seen as representative for description
		const byId: Record<string, CTWithLib[]> = {};
		for (const e of ctIndex) (byId[e.ct.id] ||= []).push(e);

		const opts: IComboBoxOption[] = [];
		for (const [ctId, entries] of Object.entries(byId)) {
			const name = entries[0].ct.name;
			const desc = entries[0].ct.description;
			const multi = counts[ctId] > 1;

			opts.push({
				key: ctId,
				text: multi ? `${name} (${counts[ctId]} locations)` : name,
				data: { description: desc, entries },
			});
		}
		return opts.sort((a, b) => a.text.localeCompare(b.text));
	}, [ctIndex]);

	// ---------- Validation ----------
	const canContinue = useMemo(() => {
		if (activeTab === 'library') {
			if (!selectedLibraryUrl) return false;
			if (showContentTypePicker) {
				const cts = ctsByLib[selectedLibraryUrl] || [];
				// If library has exactly one CT, OK without selection
				if (cts.length === 1) return true;
				// If multiple, require a CT selection
				return !!selectedCTId;
			}
			// CT picker hidden → we allow without CT
			return true;
		} else {
			// contentType tab
			if (!selectedCTId) return false;
			// Resolve to a single library if possible; otherwise require library selection (handled below)
			const libsForCT = ctIndex
				.filter((e) => e.ct.id === selectedCTId)
				.map((e) => e.lib.serverRelativeUrl);
			if (libsForCT.length === 1) return true;
			return !!selectedLibraryUrl; // user picked a specific library among the candidates
		}
	}, [activeTab, selectedLibraryUrl, selectedCTId, showContentTypePicker, ctsByLib, ctIndex]);

	// Whenever CT changes in CT-first mode, auto-resolve library if unique
	useEffect(() => {
		if (activeTab !== 'contentType' || !selectedCTId) return;
		const libsForCT = ctIndex
			.filter((e) => e.ct.id === selectedCTId)
			.map((e) => e.lib.serverRelativeUrl);
		if (libsForCT.length === 1) {
			setSelectedLibraryUrl(libsForCT[0]);
		} else if (libsForCT.length > 1) {
			// If current selected library isn't in the filtered set, clear it
			if (!libsForCT.includes(selectedLibraryUrl)) setSelectedLibraryUrl('');
		}
	}, [activeTab, selectedCTId, ctIndex, selectedLibraryUrl]);

	// In library-first, if there's only one CT for the selected library, auto-pick it
	useEffect(() => {
		if (activeTab !== 'library' || !selectedLibraryUrl || !showContentTypePicker) return;
		const cts = ctsByLib[selectedLibraryUrl] || [];
		if (cts.length === 1) setSelectedCTId(cts[0].id);
		// If multiple, preserve user's explicit choice
	}, [activeTab, selectedLibraryUrl, showContentTypePicker, ctsByLib]);

	// ---------- Submit ----------
	const handleContinue = () => {
		try {
			const libUrl = selectedLibraryUrl || (oneLibraryOnly ? libraries[0].serverRelativeUrl : '');
			if (!libUrl) return;

			let ctId: string | undefined = selectedCTId;
			if (activeTab === 'library' && showContentTypePicker) {
				const cts = ctsByLib[libUrl] || [];
				if (cts.length === 1) ctId = cts[0].id;
			}

			const choice: DestinationChoice = {
				libraryUrl: libUrl,
				contentTypeId: ctId,
				folderPath: isNonEmptyString(folderPath) ? trimLeadingSlash(folderPath!) : undefined,
			};

			onSubmit(choice);
		} catch (e) {
			setErrorMsg(normalizeError(e).message);
		}
	};

	// ---------- Render helpers ----------
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
								onChange={(_, opt) => setSelectedCTId(opt?.key as string)}
								onInputValueChange={(value) => {
									// ignore freeform
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

				{allowFolderSelection && (
					<>
						<Label>Folder (optional)</Label>
						<TextField
							placeholder="e.g., Invoices/2025"
							value={folderPath}
							onChange={(_, v) => setFolderPath(v || '')}
						/>
					</>
				)}
			</Stack>
		);
	};

	const renderContentTypeFirst = () => {
		// Determine candidate libraries for selected CT
		const libsForSelectedCT = selectedCTId
			? ctIndex.filter((e) => e.ct.id === selectedCTId).map((e) => e.lib)
			: [];

		const multipleLibs = libsForSelectedCT.length > 1;

		return (
			<Stack tokens={{ childrenGap: 12 }}>
				<Label>Content type</Label>
				<ComboBox
					allowFreeform={false}
					autoComplete="on"
					useComboBoxAsMenuWidth
					placeholder={
						allCTOptions.length > 0 ? 'Select a content type' : 'No content types available'
					}
					options={allCTOptions}
					selectedKey={selectedCTId}
					onChange={(_, opt) => setSelectedCTId(opt?.key as string)}
				/>

				{selectedCTId && (
					<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
						{(allCTOptions.find((o) => o.key === selectedCTId)?.data?.description as string) || ''}
					</Text>
				)}

				{selectedCTId && !multipleLibs && libsForSelectedCT.length === 1 && (
					<Stack>
						<Label>Destination</Label>
						<Text>
							{libTitles[libsForSelectedCT[0].serverRelativeUrl] ||
								libsForSelectedCT[0].serverRelativeUrl}
						</Text>
					</Stack>
				)}

				{selectedCTId && multipleLibs && (
					<>
						<Label>Choose a location</Label>
						<Dropdown
							options={libsForSelectedCT.map((l) => ({
								key: l.serverRelativeUrl,
								text: libTitles[l.serverRelativeUrl] || l.serverRelativeUrl,
							}))}
							selectedKey={selectedLibraryUrl || undefined}
							onChange={(_, opt) => setSelectedLibraryUrl((opt?.key as string) || '')}
							placeholder="Select a destination"
						/>
					</>
				)}

				{allowFolderSelection && (
					<>
						<Label>Folder (optional)</Label>
						<TextField
							placeholder="e.g., Invoices/2025"
							value={folderPath}
							onChange={(_, v) => setFolderPath(v || '')}
						/>
					</>
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
