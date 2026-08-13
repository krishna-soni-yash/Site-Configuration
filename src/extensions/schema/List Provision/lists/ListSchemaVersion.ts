import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
	ensureListProvision,
	ListProvisionDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.ListSchemaVersion;

const definition: ListProvisionDefinition<string> = {
	title: LIST_TITLE,
	description: "List schema version tracker",
	templateId: 100
};

export async function provisionListSchemaVersion(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionListSchemaVersion;
