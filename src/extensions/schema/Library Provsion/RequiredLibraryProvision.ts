/*eslint-disable*/
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

export const RequiredLibraryProvision = {
    ProjectDocuments: "ProjectDocuments",
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    
    const { provisionProjectDocumentsLibrary } = await import('./libraries/ProjectDocuments');
    
    await provisionProjectDocumentsLibrary(sp);
}
