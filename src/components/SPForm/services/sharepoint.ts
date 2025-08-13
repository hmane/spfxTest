// services/sharepoint.ts - Fixed critical bugs
import { SPFI, spfi } from '@pnp/sp';
import { SPFx as PnP_SPFX } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/content-types/list';
import type { IFileUploadProgressData } from '@pnp/sp/files';

export type OverwritePolicy = 'overwrite' | 'skip' | 'suffix';

export interface ContentTypeLite {
	id: string;
	name: string;
	description?: string;
	hidden?: boolean;
}

export interface SharePointService {
	getLibraryTitle(libraryUrl: string): Promise<string>;

	fileExists(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean>;

	uploadFileWithProgress(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		file: File,
		onPct: (pct: number) => void,
		overwritePolicy: OverwritePolicy,
		chunkSizeBytes?: number,
		confirmOverwrite?: (fileName: string) => Promise<boolean>,
		contentTypeId?: string // 👈 set CT immediately after upload (optional)
	): Promise<{ itemId: number; serverRelativeUrl: string; uniqueId: string }>;

	setItemContentType(
		libraryServerRelativeUrl: string,
		itemId: number,
		contentTypeId: string
	): Promise<void>;

	getLibraryContentTypes(libraryServerRelativeUrl: string): Promise<ContentTypeLite[]>;
}

// Factory: pass SPFx context (no need to pass siteUrl explicitly)
export function createSharePointService(context: any): SharePointService {
	if (!context) throw new Error('SPFx context is required to create SharePoint service');
	const sp = spfi().using(PnP_SPFX(context));
	return new PnpSharePointService(sp);
}

class PnpSharePointService implements SharePointService {
	constructor(private readonly sp: SPFI) {
		if (!sp) throw new Error('SP instance is required');
	}

	public async getLibraryTitle(libraryUrl: string): Promise<string> {
		try {
			const list = await this.sp.web.getList(libraryUrl).select('Title')();
			return list?.Title || libraryUrl.split('/').pop() || 'Library';
		} catch (e) {
			console.warn(`getLibraryTitle failed for ${libraryUrl}`, e);
			return libraryUrl.split('/').pop() || 'Library';
		}
	}

	public async fileExists(
		libraryServerRelativeUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean> {
		try {
			const folderUrl = normalizeFolderUrl(libraryServerRelativeUrl, folderPath);
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
		chunkSizeBytes: number = 2 * 1024 * 1024, // smaller chunk => earlier callbacks
		confirmOverwrite?: (fileName: string) => Promise<boolean>,
		contentTypeId?: string
	): Promise<{ itemId: number; serverRelativeUrl: string; uniqueId: string }> {
		// Resolve folder once; set up per-call existence cache to avoid repeated HEADs
		const folderUrl = normalizeFolderUrl(libraryServerRelativeUrl, folderPath);
		const folderApi = this.sp.web.getFolderByServerRelativePath(folderUrl);
		const existsCache = new Map<string, boolean>();
		const cachedExists = async (name: string): Promise<boolean> => {
			const key = `${folderUrl}|${name}`;
			if (existsCache.has(key)) return existsCache.get(key)!;
			try {
				await folderApi.files.getByUrl(name).select('UniqueId')();
				existsCache.set(key, true);
				return true;
			} catch {
				existsCache.set(key, false);
				return false;
			}
		};

		// Compute final server name per policy (using cachedExists)
		const serverFileName = await this.resolveFileNameWithCache(
			file.name,
			overwritePolicy,
			cachedExists,
			confirmOverwrite
		);

		// Upload (v4.16): progress via IFileUploadProgressData; option is overWrite
		const addRes: any = await folderApi.files.addChunked(serverFileName, file, {
			progress: (data: IFileUploadProgressData) => {
				if (typeof data?.offset === 'number' && file.size > 0) {
					const pct = Math.min(100, Math.max(0, Math.round((data.offset / file.size) * 100)));
					onPct(pct);
				}
			},
			Overwrite: overwritePolicy === 'overwrite', // Fixed: was 'Overwrite'
			...(chunkSizeBytes ? { chunkSize: chunkSizeBytes } : {}),
		});

		// Extract UniqueId/ServerRelativeUrl (return shape varies)
		const { uniqueId, serverRelativeUrl } = await this.extractFileMetadata(
			addRes,
			folderApi,
			serverFileName
		);

		// Map file GUID → list item (robust against delete/recreate name issues)
		const item = await this.sp.web.getFileById(uniqueId).getItem();
		const idSel = await item.select('Id')();
		const itemId: number = idSel?.Id as number;
		if (typeof itemId !== 'number')
			throw new Error('Could not resolve list item ID for uploaded file');

		// OPTIONAL: set ContentType right away (best effort)
		if (contentTypeId && contentTypeId.trim()) {
			try {
				await this.setItemContentType(libraryServerRelativeUrl, itemId, contentTypeId.trim());
			} catch (error) {
				// ignore; form will still open. Worst case user sees CT picker.
				console.warn('Failed to set content type immediately:', error);
			}
		}

		return { itemId, serverRelativeUrl, uniqueId };
	}

	public async setItemContentType(
		libraryServerRelativeUrl: string,
		itemId: number,
		contentTypeId: string
	): Promise<void> {
		const list = this.sp.web.getList(libraryServerRelativeUrl);
		await list.items.getById(itemId).update({ ContentTypeId: contentTypeId });
	}

	public async getLibraryContentTypes(
		libraryServerRelativeUrl: string
	): Promise<ContentTypeLite[]> {
		try {
			const cts: any = await this.sp.web
				.getList(libraryServerRelativeUrl)
				.contentTypes.select('StringId', 'Name', 'Description', 'Hidden')();

			const arr = Array.isArray(cts) ? cts : cts?.value || [];
			return arr
				.map((ct: any) => ({
					id: ct.StringId as string,
					name: ct.Name as string,
					description: ct.Description as string | undefined,
					hidden: Boolean(ct.Hidden),
				}))
				.filter((ct: ContentTypeLite) => !!ct.id && !!ct.name);
		} catch (e) {
			console.error(`getLibraryContentTypes failed for ${libraryServerRelativeUrl}`, e);
			return [];
		}
	}

	/* ------------------------- private helpers ------------------------- */

	/** Resolve final filename based on policy, using a cached exists() */
	private async resolveFileNameWithCache(
		originalName: string,
		policy: OverwritePolicy,
		exists: (name: string) => Promise<boolean>,
		confirmOverwrite?: (fileName: string) => Promise<boolean>
	): Promise<string> {
		if (policy === 'overwrite') {
			if (confirmOverwrite && (await exists(originalName))) {
				const ok = await confirmOverwrite(originalName);
				if (!ok) {
					const err: any = new Error(`User canceled overwrite for "${originalName}"`);
					err.__skip__ = true; // UI can treat as "skipped"
					throw err;
				}
			}
			return originalName;
		}

		if (policy === 'skip') {
			if (await exists(originalName)) {
				const err: any = new Error(`File "${originalName}" already exists (policy=skip)`);
				err.__skip__ = true; // UI can treat as "skipped"
				throw err;
			}
			return originalName;
		}

		// suffix
		const { name, ext } = splitNameAndExt(originalName);
		let candidate = originalName;
		let i = 0;
		while (await exists(candidate)) {
			i++;
			candidate = `${name} (${i})${ext}`;
			if (i > 200) break; // safety
		}
		return candidate;
	}

	/** Support multiple addChunked return shapes; fall back to name lookup */
	private async extractFileMetadata(
		addRes: any,
		folderApi: any,
		fileName: string
	): Promise<{ uniqueId: string; serverRelativeUrl: string }> {
		if (addRes?.UniqueId && addRes?.ServerRelativeUrl) {
			return {
				uniqueId: addRes.UniqueId as string,
				serverRelativeUrl: addRes.ServerRelativeUrl as string,
			};
		}
		if (addRes?.file) {
			const meta = await addRes.file.select('UniqueId', 'ServerRelativeUrl')();
			return {
				uniqueId: meta?.UniqueId as string,
				serverRelativeUrl: meta?.ServerRelativeUrl as string,
			};
		}
		// last resort (rare): resolve by name
		const meta = await folderApi.files.getByUrl(fileName).select('UniqueId', 'ServerRelativeUrl')();
		return {
			uniqueId: meta?.UniqueId as string,
			serverRelativeUrl: meta?.ServerRelativeUrl as string,
		};
	}
}

/* ----------------------------- utilities ---------------------------- */

function splitNameAndExt(fileName: string): { name: string; ext: string } {
	const i = fileName.lastIndexOf('.');
	if (i <= 0) return { name: fileName, ext: '' };
	return { name: fileName.substring(0, i), ext: fileName.substring(i) };
}

function normalizeFolderUrl(libraryServerRelativeUrl: string, folderPath?: string): string {
	if (!folderPath) return libraryServerRelativeUrl;
	const base = libraryServerRelativeUrl.replace(/\/+$/, '');
	const sub = folderPath.replace(/^\/+/, '');
	return `${base}/${sub}`.replace(/\/+/g, '/');
}
