// utils/links.ts

/**
 * Context-bound SharePoint link builders (pure URL helpers).
 * -------------------------------------------------------------
 * Callers pass only the file path/URL:
 *  - absolute:        "https://tenant.sharepoint.com/sites/Eng/Shared Documents/Spec.docx"
 *  - server-relative: "/sites/Eng/Shared Documents/Spec.docx"
 *  - site-relative:   "Shared Documents/Spec.docx"
 *
 * We resolve to absolute URLs using the *current web* (bind.webAbsoluteUrl / bind.webRelativeUrl).
 *
 * Bonus:
 * - OneDrive-style preview (compact WOPI preview)
 * - Async helper: version history *by path* (resolves ListId + ItemId for you)
 *
 * Notes:
 * - All WOPI/Doc.aspx routes are hosted on the current web but accept an absolute `sourcedoc`,
 *   so they can show files from other webs too (auth still applies).
 * - Keep this file focused on *link building*. No side-effects (except the async resolver).
 */

import { Context } from '../context/pnpjs-config'; // used only by linksAsync
import type { SPFI } from '@pnp/sp';

const enc = encodeURIComponent;

type Bind = {
	webAbsoluteUrl: string; // e.g., https://tenant.sharepoint.com/sites/Eng
	webRelativeUrl: string; // e.g., /sites/Eng
};

export type LinksBound = ReturnType<typeof buildLinks>;

// ---------------- public API (context-bound) ----------------

export function buildLinks(bind: Bind) {
	const webAbs = trimTrailingSlash(bind.webAbsoluteUrl);
	const webRel = ensureLeadingSlash(bind.webRelativeUrl);

	// Normalizers (keep callers simple: pass *any* style of path)
	const toAbs = (fileUrlOrPath: string) => normalizeToAbs(fileUrlOrPath, webAbs, webRel);
	const toAbsFromListRoot = (listRootUrl: string) => normalizeToAbs(listRootUrl, webAbs, webRel);

	return {
		// ---------- FILE / DOCUMENT ----------
		file: {
			/** Absolute URL from the current web + given path/URL */
			absolute: (fileUrlOrPath: string) => toAbs(fileUrlOrPath),

			/** Quick viewer (forces browser to try inline view) */
			quickView: (fileUrlOrPath: string) => `${toAbs(fileUrlOrPath)}?web=1`,

			/** Force download (bypass viewer) */
			download: (fileUrlOrPath: string) => `${toAbs(fileUrlOrPath)}?download=1`,

			/**
			 * WOPI: read-only view (best for Office/PDFs).
			 * Reliable for modern SPO; respects file permissions.
			 */
			browserView: (fileUrlOrPath: string) =>
				`${webAbs}/_layouts/15/WopiFrame.aspx?sourcedoc=${enc(toAbs(fileUrlOrPath))}&action=view`,

			/** WOPI: edit in browser (if the file type supports Office Web Apps) */
			browserEdit: (fileUrlOrPath: string) =>
				`${webAbs}/_layouts/15/WopiFrame.aspx?sourcedoc=${enc(toAbs(fileUrlOrPath))}&action=edit`,

			/** WOPI: embed view suitable for iframes */
			embedView: (fileUrlOrPath: string) =>
				`${webAbs}/_layouts/15/WopiFrame.aspx?sourcedoc=${enc(
					toAbs(fileUrlOrPath)
				)}&action=embedview`,

			/** WOPI: interactive preview (nice for PDFs/images; also used by OneDrive UI) */
			interactivePreview: (fileUrlOrPath: string) =>
				`${webAbs}/_layouts/15/WopiFrame.aspx?sourcedoc=${enc(
					toAbs(fileUrlOrPath)
				)}&action=interactivepreview`,

			/**
			 * OneDrive-style compact preview.
			 * Under the hood this is the same WOPI preview with small-view hints.
			 * Works well for quick read-only previews without full chrome.
			 */
			oneDrivePreview: (fileUrlOrPath: string) => {
				const src = enc(toAbs(fileUrlOrPath));
				// Extra flags (harmless if ignored by the viewer):
				//  - wdSmallView=1        → compact preview layout
				//  - hideComments=1       → de-clutter the UI
				//  - fromOD=true          → helpful hint for telemetry/UX (ignored by older viewers)
				return `${webAbs}/_layouts/15/WopiFrame.aspx?sourcedoc=${src}&action=interactivepreview&wdSmallView=1&hideComments=1&fromOD=true`;
			},

			/**
			 * Modern Doc.aspx route (SharePoint viewer chrome).
			 * Good when you want comments/activity pane etc.
			 */
			docPageView: (fileUrlOrPath: string) => {
				const abs = toAbs(fileUrlOrPath);
				const name = fileNameFrom(abs);
				return `${webAbs}/_layouts/15/Doc.aspx?sourcedoc=${enc(abs)}&file=${enc(
					name
				)}&action=default`;
			},

			/** Modern Doc.aspx edit (if supported) */
			docPageEdit: (fileUrlOrPath: string) => {
				const abs = toAbs(fileUrlOrPath);
				const name = fileNameFrom(abs);
				return `${webAbs}/_layouts/15/Doc.aspx?sourcedoc=${enc(abs)}&file=${enc(name)}&action=edit`;
			},

			/**
			 * Hint to open in the desktop client app (Word/Excel/PowerPoint/Visio/OneNote).
			 * Falls back to web=0 hint if scheme is unknown.
			 */
			openInClient: (fileUrlOrPath: string) => {
				const abs = toAbs(fileUrlOrPath);
				const scheme = clientSchemeFor(abs); // ms-word:ofe|u|, ms-excel:ofe|u|, etc.
				return scheme ? `${scheme}${abs}` : `${abs}?web=0`;
			},

			/**
			 * Version history page (needs ListId + ItemId).
			 * If you prefer a "by path" version, see linksAsync.versionHistoryByPath().
			 */
			versionHistory: (listId: string, itemId: number) =>
				`${webAbs}/_layouts/15/VersionHistory.aspx?list=${enc(listId)}&ID=${itemId}`,
		},

		// ---------- LIST / ITEM FORMS ----------
		listItem: {
			/**
			 * Classic display form
			 * @param listRootUrl list root path can be server-/site-relative or absolute (e.g., "/sites/Eng/Lists/Issues" or "Lists/Issues")
			 */
			display: (listRootUrl: string, id: number, source?: string) =>
				`${toAbsFromListRoot(listRootUrl)}/DispForm.aspx?ID=${id}${
					source ? `&Source=${enc(source)}` : ''
				}`,

			/** Classic edit form */
			edit: (listRootUrl: string, id: number, source?: string) =>
				`${toAbsFromListRoot(listRootUrl)}/EditForm.aspx?ID=${id}${
					source ? `&Source=${enc(source)}` : ''
				}`,

			/** Modern display by ListId + ItemId (uses current web) */
			modernDisplay: (listId: string, id: number, source?: string) =>
				`${webAbs}/_layouts/15/listform.aspx?PageType=4&ListId=${enc(listId)}&ID=${id}${
					source ? `&Source=${enc(source)}` : ''
				}`,

			/** Modern edit by ListId + ItemId */
			modernEdit: (listId: string, id: number, source?: string) =>
				`${webAbs}/_layouts/15/listform.aspx?PageType=6&ListId=${enc(listId)}&ID=${id}${
					source ? `&Source=${enc(source)}` : ''
				}`,

			/** Modern new form by ListId */
			modernNew: (listId: string, source?: string) =>
				`${webAbs}/_layouts/15/listform.aspx?PageType=8&ListId=${enc(listId)}${
					source ? `&Source=${enc(source)}` : ''
				}`,
		},

		/**
		 * Open library with upload panel.
		 * @param libraryRootUrl library path can be server-/site-relative or absolute
		 */
		uploadTo: (libraryRootUrl: string, source?: string) =>
			`${toAbsFromListRoot(libraryRootUrl)}/Forms/AllItems.aspx?upload=1${
				source ? `&Source=${enc(source)}` : ''
			}`,

		// ---------- SITE / LIST UTILITIES ----------
		site: {
			contents: () => `${webAbs}/_layouts/15/viewlsts.aspx`,
			recycleBin: () => `${webAbs}/_layouts/15/RecycleBin.aspx`,
			settings: () => `${webAbs}/_layouts/15/settings.aspx`,
		},

		list: {
			settings: (listId: string) => `${webAbs}/_layouts/15/listedit.aspx?List=${enc(listId)}`,
			permissions: (listId: string) =>
				`${webAbs}/_layouts/15/User.aspx?obj=${enc(listId)},List&List=${enc(listId)}`,
		},

		/**
		 * Escape hatch for cross-site scenarios:
		 * if you must target a different web explicitly, use LinksUnbound.*
		 */
		unbound: LinksUnbound,
	};
}

// ---------------- async helpers (no context params!) ----------------

/**
 * linksAsync
 * ----------
 * Helpers that *look up* extra info (e.g., ListId/ItemId) so callers can just pass a path.
 * Uses the already-initialized Context under the hood; no need to pass sp/context around.
 */
export const linksAsync = {
	/**
	 * Build a Version History URL from a file path/URL.
	 * Steps:
	 *   1) Normalize to server-relative path for the *current* web
	 *   2) Resolve the file's ListItem (Id) and Parent List (Id)
	 *   3) Return the standard versionHistory URL
	 */
	async versionHistoryByPath(fileUrlOrPath: string): Promise<string> {
		const { sp, webAbsoluteUrl, webRelativeUrl, links } = Context.getContext();
		const webAbs = trimTrailingSlash(webAbsoluteUrl);
		const webRel = ensureLeadingSlash(webRelativeUrl);

		const abs = normalizeToAbs(fileUrlOrPath, webAbs, webRel);
		const srvRel = serverRelativeFromAbs(abs);

		// Try to get ListItem Id + ParentList Id via a single call using select/expand
		// REST path shape: /_api/web/GetFileByServerRelativePath(decodedurl='...')/ListItemAllFields
		// Expand ParentList to get its Id.
		const li: any = await sp.web
			.getFileByServerRelativePath(srvRel)
			.listItemAllFields.select('Id', 'ParentList/Id')
			.expand('ParentList')();

		const itemId: number | undefined = li?.Id;
		const listId: string | undefined = li?.ParentList?.Id;

		if (!itemId || !listId) {
			// Fallback: use getItem() then fetch parent list id separately
			const item: any = await sp.web.getFileByServerRelativePath(srvRel).getItem();
			const id: number | undefined = (item as any)?.Id;
			const parentList: any = await (item as any).parentList?.select('Id')?.();
			if (!id || !parentList?.Id) {
				throw new Error('Could not resolve ListId/ItemId for version history.');
			}
			return links.file.versionHistory(parentList.Id, id);
		}

		return links.file.versionHistory(listId, itemId);
	},
};

// ---------------- unbound helpers (escape hatch) ----------------

export const LinksUnbound = {
	file: {
		quickView: (webUrl: string, serverRelOrAbs: string) =>
			`${toAbsUnbound(webUrl, serverRelOrAbs)}?web=1`,
		download: (webUrl: string, serverRelOrAbs: string) =>
			`${toAbsUnbound(webUrl, serverRelOrAbs)}?download=1`,
		browserView: (webUrl: string, serverRelOrAbs: string) =>
			`${trimTrailingSlash(webUrl)}/_layouts/15/WopiFrame.aspx?sourcedoc=${enc(
				toAbsUnbound(webUrl, serverRelOrAbs)
			)}&action=view`,
	},
	listItem: {
		modernDisplay: (webUrl: string, listId: string, id: number, source?: string) =>
			`${trimTrailingSlash(webUrl)}/_layouts/15/listform.aspx?PageType=4&ListId=${enc(
				listId
			)}&ID=${id}${source ? `&Source=${enc(source)}` : ''}`,
	},
};

// ---------------- internals ----------------

/** Normalize to absolute URL using current web; accepts abs/server/site-relative inputs */
function normalizeToAbs(input: string, webAbs: string, webRel: string): string {
	try {
		if (/^https?:\/\//i.test(input)) return input;
		// server-relative → base on current web origin
		if (input.startsWith('/')) return new URL(input, webAbs).toString();
		// site-relative → prepend current web relative path
		const combined = `${trimTrailingSlash(webRel)}/${input.replace(/^\//, '')}`;
		return new URL(combined, webAbs).toString();
	} catch {
		// If URL construction fails, return the input as-is (caller may still render it)
		return input;
	}
}

/** Convert absolute SPO URL to a server-relative path */
function serverRelativeFromAbs(absUrl: string): string {
	try {
		const u = new URL(absUrl);
		return u.pathname + u.search; // keep any query (rare but safe)
	} catch {
		return absUrl;
	}
}

/** Build absolute against an explicit web (for cross-site scenarios) */
function toAbsUnbound(webUrl: string, serverRelOrAbs: string): string {
	try {
		if (/^https?:\/\//i.test(serverRelOrAbs)) return serverRelOrAbs;
		return new URL(serverRelOrAbs, trimTrailingSlash(webUrl)).toString();
	} catch {
		return serverRelOrAbs;
	}
}

/** Choose a desktop client scheme by extension (Word/Excel/PowerPoint/Visio/OneNote) */
function clientSchemeFor(absUrl: string): string | null {
	const ext = fileExt(absUrl);
	switch (ext) {
		case 'doc':
		case 'docx':
		case 'dot':
		case 'rtf':
			return 'ms-word:ofe|u|';
		case 'xls':
		case 'xlsx':
		case 'xlsm':
		case 'csv':
			return 'ms-excel:ofe|u|';
		case 'ppt':
		case 'pptx':
		case 'ppsx':
			return 'ms-powerpoint:ofe|u|';
		case 'vsd':
		case 'vsdx':
			return 'ms-visio:ofe|u|';
		case 'one':
		case 'onepkg':
			return 'onenote:https://';
		default:
			return null;
	}
}

function fileExt(path: string): string {
	const q = path.split('?')[0];
	const dot = q.lastIndexOf('.');
	return dot >= 0 ? q.substring(dot + 1).toLowerCase() : '';
}

function fileNameFrom(path: string): string {
	const s = path.split('?')[0];
	const i = s.lastIndexOf('/');
	return i >= 0 ? s.substring(i + 1) : s;
}

function trimTrailingSlash(u: string) {
	return u.endsWith('/') ? u.slice(0, -1) : u;
}
function ensureLeadingSlash(p: string) {
	return p.startsWith('/') ? p : `/${p}`;
}
