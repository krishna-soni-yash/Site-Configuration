import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

function normalizeListId(listTitle: string, raw: unknown): string {
	const value = `${raw ?? ""}`.trim();
	if (!value) {
		throw new Error(`List "${listTitle}" did not return an Id.`);
	}
	return value.startsWith("{") ? value : `{${value}}`;
}

export async function fetchListId(sp: SPFI, listTitle: string): Promise<string> {
	try {
		const info = await sp.web.lists.getByTitle(listTitle).select("Id")();
		return normalizeListId(listTitle, info?.Id);
	} catch (error) {
		const details = error instanceof Error ? error.message : `${error ?? ""}`;
		throw new Error(`Unable to fetch Id for list "${listTitle}": ${details}`);
	}
}

export default fetchListId;
