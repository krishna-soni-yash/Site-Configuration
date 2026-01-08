/* eslint-disable */
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/folders";

import {
	FieldDefinition,
	addViewToList,
	ContentTypeBindingDefinition,
	EnsureListContentTypeOptions,
	ensureListContentTypes
} from "../List Provision/GenericListProvision";

type DraftVisibilitySetting = 0 | 1 | 2;

export interface DocumentLibraryViewDefinition<TViewField extends string> {
	title: string;
	fields: readonly TViewField[];
	rowLimit?: number;
	makeDefault?: boolean;
	includeLinkFileName?: boolean;
	linkFieldInternalName?: string;
	query?: string;
}

export interface FolderStructureDefinition {
	/**
	 * One or more folder paths relative to the library root (no leading slash).
	 * Example: ["Policies", "Templates/Letters"].
	 */
	folderPaths: readonly string[];
}

export interface DocumentLibraryProvisionDefinition<
	TFieldName extends string,
	TViewField extends string = TFieldName
> {
	title: string;
	description?: string;
	/** Default true to enable major versioning. */
	enableVersioning?: boolean;
	/** Keep the last N major versions (SharePoint default 500). */
	majorVersionLimit?: number;
	/** Toggle minor versioning (drafts). */
	enableMinorVersions?: boolean;
	/** Maximum number of minor versions to keep. */
	minorVersionLimit?: number;
	/** Controls who can see drafts; maps to SP's DraftVersionVisibility enum. */
	draftVisibility?: DraftVisibilitySetting;
	/** Enable the "New Folder" command. */
	enableFolderCreation?: boolean;
	/** Set to true when you need to manage content types on the library. */
	enableContentTypes?: boolean;
	contentTypes?: readonly ContentTypeBindingDefinition[];
	contentTypeOptions?: EnsureListContentTypeOptions;
	fields?: readonly FieldDefinition<TFieldName>[];
	defaultViewFields?: readonly TViewField[];
	removeFields?: readonly TFieldName[];
	views?: readonly DocumentLibraryViewDefinition<TViewField>[];
	folders?: FolderStructureDefinition;
}

async function fieldExists<TFieldName extends string>(sp: SPFI, listTitle: string, fieldName: TFieldName): Promise<boolean> {
	try {
		const field = await sp.web.lists.getByTitle(listTitle).fields.getByInternalNameOrTitle(fieldName).select("Id")();
		return !!field;
	} catch (error) {
		return false;
	}
}

async function ensureFieldRemovedFromDefaultView(sp: SPFI, listTitle: string, fieldName: string): Promise<void> {
	try {
		const list = sp.web.lists.getByTitle(listTitle);
		const defaultView = list.defaultView;
		const schemaXml = await defaultView.fields.getSchemaXml();
		if (schemaXml.includes(`Name="${fieldName}"`) || schemaXml.includes(`Name='${fieldName}'`)) {
			await defaultView.fields.remove(fieldName as any);
		}
	} catch (error) {
		console.warn(`Failed to remove field ${fieldName} from default view on library ${listTitle}:`, error);
	}
}

async function ensureDocumentLibraryFolders(sp: SPFI, listTitle: string, definition?: FolderStructureDefinition): Promise<void> {
	if (!definition || definition.folderPaths.length === 0) {
		return;
	}

	const list = sp.web.lists.getByTitle(listTitle);
	const root = await list.rootFolder.select("ServerRelativeUrl")();
	const rootUrl = String(root?.ServerRelativeUrl || "");
	if (!rootUrl) {
		return;
	}

	const ensureFolder = async (serverRelativePath: string) => {
		try {
			await sp.web.getFolderByServerRelativePath(serverRelativePath).select("UniqueId")();
		} catch (error) {
			try {
				await sp.web.folders.addUsingPath(serverRelativePath);
			} catch (createError) {
				console.warn(`Failed to create folder ${serverRelativePath}:`, createError);
			}
		}
	};

	for (const rawPath of definition.folderPaths) {
		const trimmed = rawPath.trim().replace(/^\/+/g, "").replace(/\/+$/g, "");
		if (!trimmed) {
			continue;
		}

		const parts = trimmed.split("/").filter(Boolean);
		let current = rootUrl;
		for (const part of parts) {
			current = `${current}/${part}`;
			await ensureFolder(current);
		}
	}
}

async function configureLibraryView<TViewField extends string>(
	sp: SPFI,
	listTitle: string,
	view: DocumentLibraryViewDefinition<TViewField>
): Promise<void> {
	const linkField = view.linkFieldInternalName || "LinkFilename";
	const includeLinkField = view.includeLinkFileName ?? true;
	const viewFields = [...(view.fields as readonly string[])];

	if (includeLinkField && !viewFields.includes(linkField)) {
		viewFields.unshift(linkField);
	}

	await addViewToList(sp, listTitle, {
		title: view.title,
		fields: viewFields,
		rowLimit: view.rowLimit,
		makeDefault: view.makeDefault,
		includeLinkTitle: false,
		query: view.query
	} as any);
}

async function ensureDefaultViewFields<TViewField extends string>(
	sp: SPFI,
	listTitle: string,
	defaultViewFields: readonly TViewField[]
): Promise<void> {
	if (defaultViewFields.length === 0) {
		return;
	}

	const list = sp.web.lists.getByTitle(listTitle);
	const defaultView = list.defaultView;
	const schemaXml = await defaultView.fields.getSchemaXml();
	for (const viewField of defaultViewFields) {
		if (!schemaXml.includes(`Name="${viewField}"`) && !schemaXml.includes(`Name='${viewField}'`)) {
			try {
				await defaultView.fields.add(viewField as any);
			} catch (error) {
				console.warn(`Failed to add field ${viewField} to default view on library ${listTitle}:`, error);
			}
		}
	}
}

export async function ensureDocumentLibrary<
	TFieldName extends string,
	TViewField extends string = TFieldName
>(sp: SPFI, definition: DocumentLibraryProvisionDefinition<TFieldName, TViewField>): Promise<void> {
	const {
		title,
		description = "",
		enableVersioning = true,
		majorVersionLimit,
		enableMinorVersions,
		minorVersionLimit,
		draftVisibility,
		enableFolderCreation,
		enableContentTypes,
		contentTypes = [],
		contentTypeOptions,
		fields = [],
		defaultViewFields = [],
		removeFields = [],
		views = [],
		folders
	} = definition;

	const ensureResult = await sp.web.lists.ensure(title, description, 101);
	const list = ensureResult.list;

	const updatePayload: Record<string, unknown> = {};

	if (enableVersioning !== undefined) {
		updatePayload.EnableVersioning = enableVersioning;
	}

	if (majorVersionLimit !== undefined) {
		updatePayload.MajorVersionLimit = majorVersionLimit;
	}

	if (enableMinorVersions !== undefined) {
		updatePayload.EnableMinorVersions = enableMinorVersions;
	}

	if (minorVersionLimit !== undefined) {
		updatePayload.MajorWithMinorVersionsLimit = minorVersionLimit;
	}

	if (draftVisibility !== undefined) {
		updatePayload.DraftVersionVisibility = draftVisibility;
	}

	if (enableFolderCreation !== undefined) {
		updatePayload.EnableFolderCreation = enableFolderCreation;
	}

	if (enableContentTypes !== undefined) {
		updatePayload.ContentTypesEnabled = enableContentTypes;
	}

	if (Object.keys(updatePayload).length > 0) {
		try {
			await list.update(updatePayload);
		} catch (error) {
			console.warn(`Failed to update library settings for ${title}:`, error);
		}
	}

	for (const field of fields) {
		const exists = await fieldExists(sp, title, field.internalName);
		if (!exists) {
			try {
				await list.fields.createFieldAsXml(field.schemaXml);
			} catch (error) {
				console.warn(`Failed to create field ${field.internalName} on library ${title}:`, error);
			}
		}
	}

	for (const fieldName of removeFields) {
		try {
			const exists = await fieldExists(sp, title, fieldName);
			if (exists) {
				await list.fields.getByInternalNameOrTitle(fieldName).delete();
				await ensureFieldRemovedFromDefaultView(sp, title, fieldName);
			}
		} catch (error) {
			console.warn(`Failed to remove field ${fieldName} from library ${title}:`, error);
		}
	}

	await ensureDefaultViewFields(sp, title, defaultViewFields);

	for (const view of views) {
		await configureLibraryView(sp, title, view as DocumentLibraryViewDefinition<string>);
	}

	if (contentTypes.length > 0) {
		try {
			await ensureListContentTypes(sp, title, contentTypes, contentTypeOptions);
		} catch (error) {
			console.warn(`Failed to ensure content types on library ${title}:`, error);
		}
	}

	await ensureDocumentLibraryFolders(sp, title, folders);
}
