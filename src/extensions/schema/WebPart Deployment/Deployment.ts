/*eslint-disable*/
import * as pnp from "@pnp/sp";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/clientside-pages/web";
import {
	ClientsidePageFromFile,
	CreateClientsidePage,
	ClientsideWebpart,
	IClientsidePage,
	IClientsidePageComponent,
} from "@pnp/sp/clientside-pages";
import { CheckinType } from "@pnp/sp/files";

import WebPartList, { IListBindingConfig, IWebPartEntry } from './WebPartList';

const normalizeGuid = (value?: string): string => (value || '').replace(/[{}]/g, '').toLowerCase();

export async function deployWebParts(spInstance?: SPFI): Promise<void> {
		const sp: SPFI = spInstance || (pnp as any).sp || (pnp as any).default || (pnp as unknown as SPFI);
		const webInfo = await sp.web.select("ServerRelativeUrl")();
		const webRel = webInfo?.ServerRelativeUrl || '';
	const webRelNoSlash = webRel.replace(/\/$/, '');

	const availableWebparts = await getWebpartDefinitions(sp);

	for (const entry of WebPartList) {
		const pageFileName = `${entry.pageName}.aspx`;
		const pageServerRelativeUrl = `${webRelNoSlash}/SitePages/${pageFileName}`;

		let page: IClientsidePage | undefined = await loadExistingPage(sp, pageServerRelativeUrl);
		let createdNewPage = false;

		if (!page) {
			page = await createClientsidePage(sp, pageFileName, entry.pageName);
			if (!page) {
				console.error(`Failed to ensure page ${entry.pageName}, skipping webpart add.`);
				continue;
			}
			createdNewPage = true;
		}
		const shouldSetAsHome = createdNewPage && entry.homePage === true;

		try {
			const componentDef = findWebpartDefinition(availableWebparts, entry);
			if (!componentDef) {
				console.warn(`Skipping webpart ${describeEntry(entry)}: definition not found.`);
				continue;
			}

			const alreadyHasWebpart = hasWebpart(page, componentDef);
			if (alreadyHasWebpart) {
				continue;
			}

			const webpartControl = ClientsideWebpart.fromComponentDef(componentDef);
			await configureWebpart(sp, webpartControl, entry);

			const targetColumn = ensureDefaultColumn(page);
			targetColumn.addControl(webpartControl);

			await page.save(true);
			await finalizePage(sp, pageServerRelativeUrl);

			if (shouldSetAsHome) {
				await setHomePage(sp, `SitePages/${pageFileName}`);
			}
		} catch (e) {
			console.error(`Failed to add webpart ${describeEntry(entry)} to page ${entry.pageName}:`, e);
		}
	}
}

async function loadExistingPage(sp: SPFI, pageServerRelativeUrl: string): Promise<IClientsidePage | undefined> {
	try {
		const file = sp.web.getFileByServerRelativePath(pageServerRelativeUrl);
		const page = await ClientsidePageFromFile(file);
		await page.load();
		return page;
	} catch (error) {
		return undefined;
	}
}

async function createClientsidePage(sp: SPFI, fileName: string, title: string): Promise<IClientsidePage | undefined> {
	try {
		const addResult: any = await sp.web.addClientsidePage(fileName, title);
		const page: IClientsidePage | undefined = addResult?.page || addResult;
		if (page) {
			await page.load();
			return page;
		}
		console.warn(`addClientsidePage did not return a page instance for ${fileName}.`);
		return undefined;
	} catch (error) {
		try {
			const fallbackPage = await CreateClientsidePage(sp.web, fileName, title);
			await fallbackPage.load();
			return fallbackPage;
		} catch (fallbackError) {
			console.error(`CreateClientsidePage fallback failed for ${fileName}.`, fallbackError);
			return undefined;
		}
	}
}

function ensureDefaultColumn(page: IClientsidePage) {
	const section = page.sections[0] || page.addSection();
	return section.columns[0] || section.addColumn(12);
}

function hasWebpart(page: IClientsidePage, componentDef: IClientsidePageComponent): boolean {
	const targetIds = collectDefinitionIds(componentDef);
	if (!targetIds.length) {
		return false;
	}

	const matchingControl = page.findControl(control => {
		const data: any = control.data || {};
		const candidateIds = [
			normalizeGuid(data.webPartId),
			normalizeGuid(data?.webPartData?.id),
			normalizeGuid(data?.id),
		].filter(Boolean);
		return candidateIds.some(id => targetIds.includes(id));
	});
	return !!matchingControl;
}

async function getWebpartDefinitions(sp: SPFI): Promise<IClientsidePageComponent[]> {
	try {
		return await sp.web.getClientsideWebParts();
	} catch (error) {
		return [];
	}
}

function findWebpartDefinition(defs: IClientsidePageComponent[], entry: IWebPartEntry): IClientsidePageComponent | undefined {
	const desiredIds = [normalizeGuid(entry.id), normalizeGuid((entry as any)?.componentId)].filter(Boolean);
	const desiredAliases = collectEntryAliases(entry);

	for (const def of defs) {
		const candidateIds = collectDefinitionIds(def);
		if (desiredIds.length && candidateIds.some(id => desiredIds.includes(id))) {
			return def;
		}
		if (desiredAliases.length) {
			const aliases = collectDefinitionAliases(def);
			if (aliases.some(alias => desiredAliases.includes(alias))) {
				return def;
			}
		}
	}

	return undefined;
}

async function configureWebpart(sp: SPFI, webpartControl: ClientsideWebpart, entry: IWebPartEntry): Promise<void> {
	if (entry.listBinding) {
		const listConfig = await resolveListWebpartConfig(sp, entry.listBinding);
		if (listConfig.webPartTitle) {
			webpartControl.title = listConfig.webPartTitle;
		}
		if (listConfig.properties) {
			webpartControl.setProperties(listConfig.properties);
		}
		if (listConfig.serverProcessedContent) {
			webpartControl.setServerProcessedContent(listConfig.serverProcessedContent);
		}
	}
}

function collectDefinitionIds(def: IClientsidePageComponent): string[] {
	const values = new Set<string>();
	const push = (value?: string) => {
		const normalized = normalizeGuid(value);
		if (normalized) {
			values.add(normalized);
		}
	};

	push(def.Id as string);
	const defAny: any = def;
	push(defAny?.ComponentId);
	push(defAny?.ComponentDefinition?.Id);

	const manifest = parseManifest(defAny?.Manifest);
	if (manifest) {
		push(manifest?.id);
		push(manifest?.Id);
		const componentDefinition = manifest?.componentDefinition || manifest?.ComponentDefinition;
		push(componentDefinition?.id);
		push(componentDefinition?.Id);
		const preconfigured = manifest?.preconfiguredEntries || manifest?.PreconfiguredEntries;
		if (Array.isArray(preconfigured)) {
			for (const pre of preconfigured) {
				push(pre?.webPartId);
				push(pre?.id);
				push(pre?.Id);
			}
		}
	}

	return Array.from(values);
}

function collectDefinitionAliases(def: IClientsidePageComponent): string[] {
	const aliases = new Set<string>();
	const defAny: any = def;
	const add = (value?: string) => {
		if (typeof value === 'string') {
			const trimmed = value.trim().toLowerCase();
			if (trimmed) {
				aliases.add(trimmed);
			}
		}
	};

	add(def.Name);
	add(defAny?.ComponentName);
	add(defAny?.ComponentAlias);

	const manifest = parseManifest(defAny?.Manifest);
	if (manifest) {
		add(manifest?.alias);
		add(manifest?.Alias);
		add(manifest?.componentAlias);
		add(manifest?.componentName);
		const entries = manifest?.preconfiguredEntries || manifest?.PreconfiguredEntries;
		if (Array.isArray(entries)) {
			for (const entry of entries) {
				add(toLowerValue(entry?.title));
				add(toLowerValue(entry?.Title));
				if (entry?.title && typeof entry.title === 'object') {
					for (const value of Object.values(entry.title)) {
						add(toLowerValue(`${value ?? ''}`));
					}
				}
			}
		}
	}

	return Array.from(aliases);
}

function collectEntryAliases(entry: IWebPartEntry): string[] {
	const aliases = new Set<string>();
	const add = (value?: string) => {
		if (typeof value === 'string') {
			const trimmed = value.trim().toLowerCase();
			if (trimmed) {
				aliases.add(trimmed);
			}
		}
	};

	add((entry as any)?.componentAlias);
	add(entry.alias);

	return Array.from(aliases);
}

function parseManifest(manifest: any): any | undefined {
	if (!manifest) {
		return undefined;
	}
	if (typeof manifest === 'string') {
		try {
			return JSON.parse(manifest);
		} catch (error) {
			console.warn('Unable to parse webpart manifest string.', error);
			return undefined;
		}
	}
	if (typeof manifest === 'object') {
		return manifest;
	}
	return undefined;
}

function toLowerValue(value?: string): string | undefined {
	if (typeof value !== 'string') {
		return undefined;
	}
	const trimmed = value.trim();
	return trimmed ? trimmed.toLowerCase() : undefined;
}

function describeEntry(entry: IWebPartEntry): string {
	return entry.alias || entry.id || entry.pageName;
}

interface IListWebpartConfigResult {
	properties: Record<string, any>;
	serverProcessedContent?: Record<string, any>;
	webPartTitle?: string;
}

async function resolveListWebpartConfig(sp: SPFI, binding: IListBindingConfig): Promise<IListWebpartConfigResult> {
	const listSelector = binding.listId ? sp.web.lists.getById(stripGuid(binding.listId)) : sp.web.lists.getByTitle(binding.listTitle);
	const listInfo: any = await listSelector
		.select('Id', 'Title', 'DefaultViewUrl', 'RootFolder/ServerRelativeUrl', 'BaseTemplate')
		.expand('RootFolder')();
	const listGuid = stripGuid(binding.listId || listInfo?.Id);
	const listTitle = binding.listTitle || listInfo?.Title;
	const listUrl = deriveListUrl(listInfo);
	const baseTemplateRaw = listInfo?.BaseTemplate;
	const baseTemplate = typeof baseTemplateRaw === 'number'
		? baseTemplateRaw
		: parseInt(`${baseTemplateRaw ?? ''}`, 10);
	const isDocumentLibrary = binding.isDocumentLibrary !== undefined
		? binding.isDocumentLibrary
		: baseTemplate === 101;

	let viewGuid = stripGuid(binding.viewId);
	let viewTitle = binding.viewTitle;
	let viewUrl: string | undefined;

	if (!viewGuid) {
		if (binding.viewTitle) {
			const viewInfo: any = await listSelector.views.getByTitle(binding.viewTitle).select('Id', 'Title', 'ServerRelativeUrl')();
			viewGuid = stripGuid(viewInfo?.Id);
			viewTitle = viewInfo?.Title || binding.viewTitle;
			viewUrl = viewInfo?.ServerRelativeUrl;
		} else {
			const viewInfo: any = await listSelector.defaultView.select('Id', 'Title', 'ServerRelativeUrl')();
			viewGuid = stripGuid(viewInfo?.Id);
			viewTitle = viewInfo?.Title || viewTitle;
			viewUrl = viewInfo?.ServerRelativeUrl;
		}
	}

	const displayTitle = binding.webPartTitle || viewTitle || listTitle;

	const properties: Record<string, any> = {
		isDocumentLibrary,
		listId: listGuid,
		listTitle,
		listUrl,
		selectedListId: listGuid,
		selectedListTitle: listTitle,
		selectedListUrl: listUrl,
		selectedViewId: viewGuid,
		selectedViewTitle: viewTitle,
		selectedViewUrl: viewUrl,
	};

	return {
		properties,
		serverProcessedContent: {
			searchablePlainTexts: {
				title: displayTitle || '',
			},
		},
		webPartTitle: displayTitle,
	};
}

function stripGuid(value?: string): string {
	const raw = `${value ?? ''}`.trim();
	return raw ? raw.replace(/[{}]/g, '').toLowerCase() : '';
}

function deriveListUrl(listInfo: any): string {
	const rootFolderUrl = listInfo?.RootFolder?.ServerRelativeUrl;
	if (typeof rootFolderUrl === 'string' && rootFolderUrl.trim()) {
		return rootFolderUrl;
	}
	const defaultViewUrl = `${listInfo?.DefaultViewUrl ?? ''}`;
	return defaultViewUrl.replace(/\/Forms\/.+$/i, '');
}

async function finalizePage(sp: SPFI, pageServerRelativeUrl: string): Promise<void> {
	const file = sp.web.getFileByServerRelativePath(pageServerRelativeUrl);

	try {
		await file.checkin("Automated check-in", CheckinType.Major);
	} catch (error: any) {
		const message = (error?.message || "").toString();
		if (!message.includes("is not checked out")) {
			console.warn(`checkIn failed for ${pageServerRelativeUrl}`, error);
		}
	}

	try {
		await file.publish("Automated publish");
	} catch (error: any) {
		const message = (error?.message || "").toString();
		if (!message.includes("has not been checked out")) {
			console.warn(`publish failed for ${pageServerRelativeUrl}`, error);
		}
	}
}

async function setHomePage(sp: SPFI, welcomePage: string): Promise<void> {
	try {
		await sp.web.rootFolder.update({ WelcomePage: welcomePage });
	} catch (error) {
		console.error(`Failed to set ${welcomePage} as the home page.`, error);
	}
}

export default deployWebParts;