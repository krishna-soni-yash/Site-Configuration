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

const LIST_TITLE = RequiredListsProvision.OpenIssues;

type OpenIssuesFieldName =
    | "OpenIssuesCount";

type OpenIssuesViewField = OpenIssuesFieldName;

const fieldDefinitions: readonly FieldDefinition<OpenIssuesFieldName>[] = [
    {
        internalName: "OpenIssuesCount",
        schemaXml: `<Field Type='Number' Name='OpenIssuesCount' StaticName='OpenIssuesCount' DisplayName='OpenIssuesCount' Decimals='2' />`
    }
] as const;

const defaultViewFields: readonly OpenIssuesViewField[] = [
    "OpenIssuesCount"
] as const;

const definition: ListProvisionDefinition<OpenIssuesFieldName, OpenIssuesViewField> = {
    title: LIST_TITLE,
    description: "Open Issues",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionOpenIssues(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionOpenIssues;
