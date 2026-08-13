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

const LIST_TITLE = RequiredListsProvision.Velocity;

type VelocityFieldName =
    | "ApplicableGraphID"
    | "Goal"
    | "USL"
    | "LSL"
    | "ProjectType"
    | "Velocity"
    | "TotalDeliveredStoryPoints"
    | "TeamSize";

type VelocityViewField = VelocityFieldName;

const fieldDefinitions: readonly FieldDefinition<VelocityFieldName>[] = [
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
        internalName: "Velocity",
        schemaXml: `<Field Type='Number' Name='Velocity' StaticName='Velocity' DisplayName='Velocity' Decimals='2' />`
    },
    {
        internalName: "TotalDeliveredStoryPoints",
        schemaXml: `<Field Type='Number' Name='TotalDeliveredStoryPoints' StaticName='TotalDeliveredStoryPoints' DisplayName='TotalDeliveredStoryPoints' Decimals='2' />`
    },
    {
        internalName: "TeamSize",
        schemaXml: `<Field Type='Number' Name='TeamSize' StaticName='TeamSize' DisplayName='TeamSize' Decimals='2' />`
    }
] as const;

const defaultViewFields: readonly VelocityFieldName[] = [
    "ApplicableGraphID",
    "Goal",
    "USL",
    "LSL",
    "ProjectType",
    "Velocity",
    "TotalDeliveredStoryPoints",
    "TeamSize"
] as const;

const definition: ListProvisionDefinition<VelocityFieldName, VelocityViewField> = {
    title: LIST_TITLE,
    description: "",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionVelocity(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
    const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
    await ensureListProvision(sp, {
        ...definition,
        lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
    });
}

export default provisionVelocity;