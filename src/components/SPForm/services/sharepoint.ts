// src/webparts/UploadAndEdit/services/sharepoint.ts

import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/items';
import type { IFileUploadProgressData } from '@pnp/sp/files';

export type OverwritePolicy = 'overwrite' | 'skip' | 'suffix';

export type ContentTypeLite = {
	id: string; // StringId
	name: string; // Name
};

export interface SharePointService {
	/** Check if a file exists in (library + optional subfolder) */
	fileExists(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean>;

	/**
	 * Upload a file with progress + overwrite policy handling.
	 * - Uses chunked upload under the hood
	 * - Returns created list item id (resolved via file GUID)
	 */
	uploadFileWithProgress(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		file: File,
		onPct: (pct: number) => void,
		overwritePolicy: OverwritePolicy,
		chunkSizeBytes?: number,
		confirmOverwrite?: (fileName: string) => Promise<boolean>
	): Promise<{ itemId: number; serverRelativeUrl: string; uniqueId: string }>;

	/** Force an item’s content type */
	setItemContentType(
		libraryServerRelativeUrl: string,
		itemId: number,
		contentTypeId: string
	): Promise<void>;

	/** Get content types available for a library (optionally filtered) */
	getLibraryContentTypes(libraryServerRelativeUrl: string): Promise<ContentTypeLite[]>;
}

export function createSharePointService(siteUrl: string, spfxContext: any): SharePointService {
	const sp = spfi(siteUrl).using(SPFx(spfxContext));
	return new PnpSharePointService(sp);
}

/* --------------------------------- Impl --------------------------------- */

class PnpSharePointService implements SharePointService {
	constructor(private readonly sp: SPFI) {}

	/**
	 * Fast existence check: HEAD-like request by selecting UniqueId.
	 * Uses a tiny per-call cache by (folderUrl|name) to avoid duplicates in a batch.
	 */
	public async fileExists(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean> {
		const folderUrl = normalizeFolderUrl(libraryServerRelativeUrl, folderPath);
		try {
			await this.sp.web
				.getFolderByServerRelativePath(folderUrl)
				.files.getByUrl(fileName)
				.select('UniqueId')();
			return true;
		} catch {
			return false;
		}
	}

	public async uploadFileWithProgress(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		file: File,
		onPct: (pct: number) => void,
		overwritePolicy: OverwritePolicy,
		chunkSizeBytes?: number,
		confirmOverwrite?: (fileName: string) => Promise<boolean>
	): Promise<{ itemId: number; serverRelativeUrl: string; uniqueId: string }> {
		// ---- resolve folder once & set up per-call caches ----
		const folderUrl = normalizeFolderUrl(libraryServerRelativeUrl, folderPath);
		const folderApi = this.sp.web.getFolderByServerRelativePath(folderUrl);

		// per-call existence cache to avoid re-checking the same file multiple times
		const existsCache = new Map<string, boolean>();
		const cachedExists = async (name: string): Promise<boolean> => {
			const key = `${folderUrl}|${name}`;
			if (existsCache.has(key)) return existsCache.get(key)!;
			const yes = await this.fileExists(libraryServerRelativeUrl, folderPath, name);
			existsCache.set(key, yes);
			return yes;
		};

		// ---- compute target server file name based on policy ----
		let serverFileName = file.name;

		if (overwritePolicy === 'suffix') {
			const { name, ext } = splitNameAndExt(serverFileName);
			let i = 0;
			while (await cachedExists(serverFileName)) {
				i++;
				serverFileName = `${name} (${i})${ext}`;
				if (i > 100) break; // safety
			}
		} else if (overwritePolicy === 'skip') {
			if (await cachedExists(serverFileName)) {
				// Return a typed error that caller can interpret as "skipped"
				const err: any = new Error(`File "${serverFileName}" already exists (policy=skip).`);
				err.__skip__ = true;
				throw err;
			}
		} else if (overwritePolicy === 'overwrite') {
			if ((await cachedExists(serverFileName)) && confirmOverwrite) {
				const ok = await confirmOverwrite(serverFileName);
				if (!ok) {
					const err: any = new Error(`User canceled overwrite for "${serverFileName}".`);
					err.__skip__ = true;
					throw err;
				}
			}
		}

		// ---- upload (chunked) with **progress** ----
		// v4.16 option: use "overWrite" (not "Overwrite")
		const addRes: any = await folderApi.files.addChunked(serverFileName, file, {
			progress: (data: IFileUploadProgressData) => {
				if (typeof data?.offset === 'number' && file.size > 0) {
					const pct = Math.min(100, Math.max(0, Math.round((data.offset / file.size) * 100)));
					onPct(pct);
				}
			},
			overWrite: overwritePolicy === 'overwrite',
			...(chunkSizeBytes ? ({ chunkSize: chunkSizeBytes } as any) : null),
		});

		// ---- get fresh metadata; v4.16 may return metadata directly or under .file ----
		let uniqueId: string | undefined;
		let serverRelativeUrl: string | undefined;

		if (addRes?.UniqueId) {
			uniqueId = addRes.UniqueId as string;
			serverRelativeUrl = addRes.ServerRelativeUrl as string;
		} else if (addRes?.file) {
			const meta = await addRes.file.select('UniqueId', 'ServerRelativeUrl')();
			uniqueId = meta?.UniqueId as string;
			serverRelativeUrl = meta?.ServerRelativeUrl as string;
		} else {
			// last-resort: resolve by name (can be fooled by recycle bin in rare cases)
			const meta = await folderApi.files
				.getByUrl(serverFileName)
				.select('UniqueId', 'ServerRelativeUrl')();
			uniqueId = meta?.UniqueId as string;
			serverRelativeUrl = meta?.ServerRelativeUrl as string;
		}

		if (!uniqueId || !serverRelativeUrl) {
			throw new Error('Upload succeeded but could not resolve file metadata.');
		}

		// ---- map file GUID → list item (robust, avoids deleted/recreated name confusion) ----
		const item = await this.sp.web.getFileById(uniqueId).getItem();
		const info = await item.select('Id')();
		const itemId: number = info?.Id as number;

		if (typeof itemId !== 'number') {
			throw new Error('Could not resolve list item for uploaded file.');
		}

		return { itemId, serverRelativeUrl, uniqueId };
	}

	public async setItemContentType(
		libraryServerRelativeUrl: string,
		itemId: number,
		contentTypeId: string
	): Promise<void> {
		const list = this.sp.web.getList(libraryServerRelativeUrl);
		await list.items.getById(itemId).update({
			ContentTypeId: contentTypeId,
		});
	}

	public async getLibraryContentTypes(
		libraryServerRelativeUrl: string
	): Promise<ContentTypeLite[]> {
		const list: any = await this.sp.web
			.getList(libraryServerRelativeUrl)
			.contentTypes.select('StringId', 'Name')();
		// PnP returns an array-like; normalize to ContentTypeLite[]
		return (Array.isArray(list) ? list : list?.value || [])
			.map((ct: any) => ({ id: ct.StringId as string, name: ct.Name as string }))
			.filter((ct: ContentTypeLite) => !!ct.id && !!ct.name);
	}
}

/* ----------------------------- helpers ----------------------------- */

function splitNameAndExt(fileName: string): { name: string; ext: string } {
	const idx = fileName.lastIndexOf('.');
	if (idx <= 0) return { name: fileName, ext: '' };
	return { name: fileName.substring(0, idx), ext: fileName.substring(idx) };
}

function normalizeFolderUrl(libraryServerRelativeUrl: string, folderPath?: string): string {
	if (!folderPath) return libraryServerRelativeUrl;
	// Safe join: "/sites/x/Lib" + "My/Folder" => "/sites/x/Lib/My/Folder"
	const joined = `${stripTrailingSlash(libraryServerRelativeUrl)}/${stripLeadingSlash(folderPath)}`;
	return joined.replace(/\/+/g, '/');
}
function stripLeadingSlash(s: string) {
	return s?.startsWith('/') ? s.slice(1) : s;
}
function stripTrailingSlash(s: string) {
	return s?.endsWith('/') ? s.slice(0, -1) : s;
}
