/*eslint-disable*/
import { spfi, SPFx, SPFI } from "@pnp/sp";

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}