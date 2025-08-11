// src/webparts/UploadAndEdit/services/sharepoint.ts

import { SPFx as PnP_SPFX, spfi, SPFI } from '@pnp/sp';

// Bring in SP features (extends prototypes & typings)
import '@pnp/sp/content-types/list'; // list.contentTypes()
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/items';
import '@pnp/sp/lists';
import '@pnp/sp/webs';

import type { IFileUploadProgressData } from '@pnp/sp/files';

import { ContentTypeInfo, OverwritePolicy, SharePointService } from '../types';

import { encodePathSegments, safeJoinPath, trimTrailingSlash, withNumericSuffix } from '../utils';

/**
 * Factory: create a service bound to the current site/context.
 */
export function createSharePointService(siteUrl: string, spfxContext: any): SharePointServiceImpl {
	const sp = spfi(siteUrl).using(PnP_SPFX(spfxContext));
	return new SharePointServiceImpl(sp);
}

export class SharePointServiceImpl implements SharePointService {
	constructor(private readonly sp: SPFI) {}

	/**
	 * Ensure (create if needed) nested folder path inside a document library.
	 * @param libraryUrl server-relative, e.g. "/sites/Contoso/Shared Documents"
	 * @param folderPath relative to library root, e.g. "Invoices/2025"
	 */
	async ensureFolder(libraryUrl: string, folderPath?: string): Promise<void> {
		if (!folderPath) return;

		const cleanLib = trimTrailingSlash(libraryUrl);
		const segments = folderPath
			.split('/')
			.map((s) => s.trim())
			.filter(Boolean);
		let current = cleanLib;

		for (const seg of segments) {
			current = `${current}/${seg}`;
			try {
				await this.sp.web
					.getFolderByServerRelativePath(encodePathSegments(current))
					.select('ServerRelativeUrl')();
			} catch {
				const parent = current.substring(0, current.lastIndexOf('/'));
				await this.sp.web
					.getFolderByServerRelativePath(encodePathSegments(parent))
					.folders.addUsingPath(seg);
			}
		}
	}

	/**
	 * Get displayable content types for a library (filters hidden).
	 */
	async getLibraryContentTypes(libraryUrl: string): Promise<ContentTypeInfo[]> {
		const raw: any[] = await this.sp.web
			.getList(libraryUrl)
			.contentTypes.select('Name', 'Id', 'StringId', 'Description', 'Hidden', 'Sealed', 'Group')();

		const mapped: ContentTypeInfo[] = raw.map((ct: any) => ({
			id: ct?.StringId ?? ct?.Id?.StringValue ?? ct?.Id ?? '',
			name: ct?.Name ?? '',
			description: ct?.Description ?? '',
			hidden: !!ct?.Hidden,
			sealed: !!ct?.Sealed,
			group: ct?.Group ?? '',
		}));

		return mapped.filter((ct) => !ct.hidden);
	}

	/**
	 * Resolve library Title for display.
	 */
	async getLibraryTitle(libraryUrl: string): Promise<string> {
		const list: any = await this.sp.web.getList(libraryUrl).select('Title')();
		return (list?.Title as string) ?? libraryUrl.split('/').pop() ?? libraryUrl;
	}

	/**
	 * Chunked upload with progress (PnPjs v3).
	 * Honors overwrite policy: overwrite | skip | suffix.
	 * Returns the created/updated item Id.
	 */
	async uploadFileWithProgress(
		libraryUrl: string,
		folderPath: string | undefined,
		file: File,
		onProgress: (pct: number) => void,
		overwritePolicy: OverwritePolicy
	): Promise<{ itemId: number }> {
		const folder = await this.getTargetFolder(libraryUrl, folderPath);
		const folderApi = this.sp.web.getFolderByServerRelativePath(
			encodePathSegments(folder.serverRelativeUrl)
		);

		// Name + existence
		let serverFileName = file.name;
		const exists = await this.fileExists(libraryUrl, folderPath, serverFileName);

		if (exists) {
			if (overwritePolicy === 'skip') {
				throw new Error(`File already exists: ${serverFileName}`);
			}
			if (overwritePolicy === 'suffix') {
				serverFileName = await this.findSuffixName(folder.serverRelativeUrl, serverFileName);
			}
			// overwrite -> handled by Overwrite flag below
		}

		// New v3 signature: addChunked(name, file, { progress, Overwrite, chunkSize? })
		const result: any = await folderApi.files.addChunked(serverFileName, file, {
			progress: (data: IFileUploadProgressData) => {
				// IFileUploadProgressData: { uploadId, stage, offset }
				if (typeof data?.offset === 'number' && file.size > 0) {
					const pct = Math.min(100, Math.max(0, Math.round((data.offset / file.size) * 100)));
					onProgress(pct);
				}
			},
			Overwrite: overwritePolicy === 'overwrite',
			// chunkSize?: number // keep default (10 MB)
		});

		const item = await result.file.getItem();
		const info: any = await item.select('Id')();
		onProgress(100);
		return { itemId: info?.Id as number };
	}

	/**
	 * Set a specific Content Type on an existing item (single item).
	 */
	async setItemContentType(
		libraryUrl: string,
		itemId: number,
		contentTypeId: string
	): Promise<void> {
		await this.sp.web
			.getList(libraryUrl)
			.items.getById(itemId)
			.validateUpdateListItem([{ FieldName: 'ContentTypeId', FieldValue: contentTypeId }]);
	}

	/**
	 * Public: check if a file name exists at a destination (used by UI preflight).
	 */
	async fileExists(
		libraryUrl: string,
		folderPath: string | undefined,
		fileName: string
	): Promise<boolean> {
		const folder = await this.getTargetFolder(libraryUrl, folderPath);
		try {
			await this.sp.web
				.getFolderByServerRelativePath(encodePathSegments(folder.serverRelativeUrl))
				.files.getByUrl(encodePathSegments(fileName))
				.select('Name')();
			return true;
		} catch {
			return false;
		}
	}

	// ---------- Internals ----------

	private async getTargetFolder(
		libraryUrl: string,
		folderPath?: string
	): Promise<{ serverRelativeUrl: string }> {
		const cleanLib = trimTrailingSlash(libraryUrl);
		if (!folderPath) {
			// Validate root exists
			await this.sp.web
				.getFolderByServerRelativePath(encodePathSegments(cleanLib))
				.select('ServerRelativeUrl')();
			return { serverRelativeUrl: cleanLib };
		}
		await this.ensureFolder(cleanLib, folderPath);
		const full = safeJoinPath(cleanLib, folderPath);
		return { serverRelativeUrl: full };
	}

	private async findSuffixName(folderServerUrl: string, originalName: string): Promise<string> {
		for (let i = 1; i < 1000; i++) {
			const candidate = withNumericSuffix(originalName, i);
			if (!(await this.fileExists(folderServerUrl, undefined, candidate))) return candidate;
		}
		throw new Error('Could not find an available file name after 999 attempts.');
	}
}
