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

import WebPartList from './WebPartList';

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

		if (!page) {
			page = await createClientsidePage(sp, pageFileName, entry.pageName);
			if (!page) {
				console.error(`Failed to ensure page ${entry.pageName}, skipping webpart add.`);
				continue;
			}
		}

		try {
			const alreadyHasWebpart = hasWebpart(page, entry.id);
			if (alreadyHasWebpart) {
				console.log(`Page ${entry.pageName} already contains webpart ${entry.id}, skipping add.`);
				continue;
			}

			const componentDef = findWebpartDefinition(availableWebparts, entry.id);
			if (!componentDef) {
				console.warn(`Webpart definition not found for id ${entry.id}, skipping.`);
				continue;
			}

			const webpartControl = ClientsideWebpart.fromComponentDef(componentDef);
			const targetColumn = ensureDefaultColumn(page);
			targetColumn.addControl(webpartControl);

			await page.save(true);
			await finalizePage(sp, pageServerRelativeUrl);

			console.log(`Ensured webpart ${entry.id} on page ${entry.pageName}`);
		} catch (e) {
			console.error(`Failed to add webpart ${entry.id} to page ${entry.pageName}:`, e);
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
		// likely means the page does not exist yet
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
		console.warn(`addClientsidePage failed for ${fileName}, attempting CreateClientsidePage fallback.`, error);
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

function hasWebpart(page: IClientsidePage, componentId: string): boolean {
	const targetId = normalizeGuid(componentId);
	const matchingControl = page.findControl(control => {
		const data: any = control.data || {};
		const candidateIds = [
			normalizeGuid(data.webPartId),
			normalizeGuid(data?.webPartData?.id),
			normalizeGuid(data?.id),
		].filter(Boolean);
		return candidateIds.some(id => id === targetId);
	});
	return !!matchingControl;
}

async function getWebpartDefinitions(sp: SPFI): Promise<IClientsidePageComponent[]> {
	try {
		return await sp.web.getClientsideWebParts();
	} catch (error) {
		console.warn('Failed to retrieve clientsides webpart definitions.', error);
		return [];
	}
}

function findWebpartDefinition(defs: IClientsidePageComponent[], componentId: string): IClientsidePageComponent | undefined {
	const targetId = normalizeGuid(componentId);
	return defs.find(def => {
		const candidateIds = [
			normalizeGuid(def.Id as string),
			normalizeGuid((def as any)?.Id?.toString?.()),
			normalizeGuid((def as any)?.ComponentDefinition?.Id),
			normalizeGuid((def as any)?.Manifest?.Id),
			normalizeGuid((def as any)?.PreconfiguredEntries?.[0]?.webPartId),
		].filter(Boolean);
		return candidateIds.some(id => id === targetId);
	});
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

export default deployWebParts;