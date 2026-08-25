import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
    ensureListProvision,
    ListProvisionDefinition
} from "../List Provision/GenericListProvision";

export const CurrentWebPartPageVersion = 2;
export const WebPartPageVersionListTitle = "WebPartPageVersion";

const definition: ListProvisionDefinition<string> = {
    title: WebPartPageVersionListTitle,
    description: "Web part page schema version tracker",
    templateId: 100
};

export async function ensureWebPartPageVersionList(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export async function getWebPartPageVersion(sp: SPFI): Promise<{
    itemId?: number;
    version: number;
}> {
    const items = await sp.web.lists
        .getByTitle(WebPartPageVersionListTitle)
        .items
        .select("ID", "Title")
        .top(5000)();
    const latest = items.reduce<{ itemId?: number; version: number }>((current, item) => {
        const parsedVersion = Number(`${item?.Title ?? ""}`.trim());
        if (Number.isFinite(parsedVersion) && parsedVersion > current.version) {
            return { itemId: item?.ID, version: parsedVersion };
        }
        return current;
    }, { version: 0 });

    return {
        itemId: latest.itemId,
        version: latest.version
    };
}

export async function recordWebPartPageVersion(
    sp: SPFI,
    itemId?: number
): Promise<void> {
    const list = sp.web.lists.getByTitle(WebPartPageVersionListTitle);

    if (itemId !== undefined) {
        await list.items.getById(itemId).update({
            Title: `${CurrentWebPartPageVersion}`
        });
        return;
    }

    await list.items.add({
        Title: `${CurrentWebPartPageVersion}`
    });
}
