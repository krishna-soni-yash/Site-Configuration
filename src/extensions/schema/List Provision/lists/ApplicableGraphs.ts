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

const LIST_TITLE = RequiredListsProvision.ApplicableGraphs;

type ApplicableGraphsFieldName =
    | "Status"
    | "UpdateFrequency";

type ApplicableGraphsViewField = ApplicableGraphsFieldName;

const fieldDefinitions: readonly FieldDefinition<ApplicableGraphsFieldName>[] = [
    {
        internalName: "Status",
        schemaXml: `<Field Type='Text' Name='Status' StaticName='Status' DisplayName='Status' />`
    },
    {
        internalName: "UpdateFrequency",
        schemaXml: `<Field Type='Text' Name='UpdateFrequency' StaticName='UpdateFrequency' DisplayName='UpdateFrequency' />`
    }
] as const;

const defaultViewFields: readonly ApplicableGraphsViewField[] = [
    "Status",
    "UpdateFrequency"
] as const;

const definition: ListProvisionDefinition<ApplicableGraphsFieldName, ApplicableGraphsViewField> = {
    title: LIST_TITLE,
    description: "Open Action Items",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionApplicableGraphs(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionApplicableGraphs;
