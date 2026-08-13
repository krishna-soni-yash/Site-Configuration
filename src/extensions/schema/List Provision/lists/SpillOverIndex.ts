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

const LIST_TITLE = RequiredListsProvision.SpillOverIndex;

type SpillOverIndexFieldName =
    | "ApplicableGraphID"
    | "Goal"
    | "USL"
    | "LSL"
    | "ProjectType"
    | "SpillOverIndex"
    | "UserStoriesDelivered"
    | "UserStoriesCommitted";

type SpillOverIndexViewField = SpillOverIndexFieldName;

const fieldDefinitions: readonly FieldDefinition<SpillOverIndexFieldName>[] = [
    {
        internalName: "Goal",
        schemaXml: `<Field Type='Text' Name='Goal' StaticName='Goal' DisplayName='Goal' MaxLength='255' />`
    },
    {
        internalName: "USL",
        schemaXml: `<Field Type='Number' Name='USL' StaticName='USL' DisplayName='USL' Decimals='2' />`
    },
    {
        internalName: "LSL",
        schemaXml: `<Field Type='Number' Name='LSL' StaticName='LSL' DisplayName='LSL' Decimals='2' />`
    },
    {
        internalName: "ProjectType",
        schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
    },
    {
        internalName: "SpillOverIndex",
        schemaXml: `<Field Type='Number' Name='SpillOverIndex' StaticName='SpillOverIndex' DisplayName='SpillOverIndex' Decimals='2' />`
    },
    {
        internalName: "UserStoriesDelivered",
        schemaXml: `<Field Type='Number' Name='UserStoriesDelivered' StaticName='UserStoriesDelivered' DisplayName='UserStoriesDelivered' Decimals='2' />`
    },
    {
        internalName: "UserStoriesCommitted",
        schemaXml: `<Field Type='Number' Name='UserStoriesCommitted' StaticName='UserStoriesCommitted' DisplayName='UserStoriesCommitted' Decimals='2' />`
    }
] as const;

const defaultViewFields: readonly SpillOverIndexFieldName[] = [
    "ApplicableGraphID",
    "Goal",
    "USL",
    "LSL",
    "ProjectType",
    "SpillOverIndex",
    "UserStoriesDelivered",
    "UserStoriesCommitted"
] as const;

const definition: ListProvisionDefinition<SpillOverIndexFieldName, SpillOverIndexViewField> = {
    title: LIST_TITLE,
    description: "",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionSpillOverIndex(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
    const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
    await ensureListProvision(sp, {
        ...definition,
        lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
    });
}

export default provisionSpillOverIndex;