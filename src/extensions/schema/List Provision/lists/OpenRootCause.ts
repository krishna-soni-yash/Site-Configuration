import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    ensureListProvision,
    FieldDefinition,
    ListProvisionDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.OpenRootCause;

type OpenRootCauseFieldName =
    | "RootCauseCount";

type OpenRootCauseViewField = OpenRootCauseFieldName;

const fieldDefinitions: readonly FieldDefinition<OpenRootCauseFieldName>[] = [
    {
        internalName: "RootCauseCount",
        schemaXml: `<Field Type='Number' Name='RootCauseCount' StaticName='RootCauseCount' DisplayName='RootCauseCount' Decimals='2' />`
    }
] as const;

const defaultViewFields: readonly OpenRootCauseViewField[] = [
    "RootCauseCount"
] as const;

const definition: ListProvisionDefinition<OpenRootCauseFieldName, OpenRootCauseViewField> = {
    title: LIST_TITLE,
    description: "Open Root Cause",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionOpenRootCause(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionOpenRootCause;
