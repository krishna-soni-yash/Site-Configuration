import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    ensureListProvision,
    FieldDefinition,
    ListProvisionDefinition,
    createLookupFieldDefinition,
    fetchListId
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.OpenActionItems;

type OpenActionItemsFieldName =
    | "ApplicableGraphID"
    | "StatusOpenCount";

type OpenActionItemsViewField = OpenActionItemsFieldName;

const fieldDefinitions: readonly FieldDefinition<OpenActionItemsFieldName>[] = [
    {
        internalName: "StatusOpenCount",
        schemaXml: `<Field Type='Number' Name='StatusOpenCount' StaticName='StatusOpenCount' DisplayName='StatusOpenCount' Decimals='2' />`
    }
] as const;

const defaultViewFields: readonly OpenActionItemsViewField[] = [
    "ApplicableGraphID",
    "StatusOpenCount"
] as const;

const definition: ListProvisionDefinition<OpenActionItemsFieldName, OpenActionItemsViewField> = {
    title: LIST_TITLE,
    description: "Open Action Items",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionOpenActionItems(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
    const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
    await ensureListProvision(sp, {
        ...definition,
        lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
    });
}

export default provisionOpenActionItems;
