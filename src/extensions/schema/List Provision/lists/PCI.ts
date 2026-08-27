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

const LIST_TITLE = RequiredListsProvision.PCI;

type PCIFieldName =
    | "ApplicableGraphID"
    | "ActualEndDate"
    | "PCI"
    | "NCCount"
    | "ObservationCount"
    | "FCFindingCount";

type PCIViewField = PCIFieldName;

const fieldDefinitions: readonly FieldDefinition<PCIFieldName>[] = [
    {
        internalName: "PCI",
        schemaXml: `<Field Type='Number' Name='PCI' StaticName='PCI' DisplayName='PCI' Decimals='2' />`
    },
    {
        internalName: "NCCount",
        schemaXml: `<Field Type='Number' Name='NCCount' StaticName='NCCount' DisplayName='NCCount' Decimals='2' />`
    },
    {
        internalName: "ObservationCount",
        schemaXml: `<Field Type='Number' Name='ObservationCount' StaticName='ObservationCount' DisplayName='ObservationCount' Decimals='2' />`
    },
    {
        internalName: "FCFindingCount",
        schemaXml: `<Field Type='Number' Name='FCFindingCount' StaticName='FCFindingCount' DisplayName='FCFindingCount' Decimals='2' />`
    },
    {
        internalName: "ActualEndDate",
        schemaXml: `<Field Type='DateTime' Name='ActualEndDate' StaticName='ActualEndDate' DisplayName='ActualEndDate' Format='DateOnly' />`
    }
] as const;

const defaultViewFields: readonly PCIFieldName[] = [
    "ApplicableGraphID",
    "ActualEndDate",
    "PCI",
    "NCCount",
    "ObservationCount",
    "FCFindingCount"
] as const;

const definition: ListProvisionDefinition<PCIFieldName, PCIViewField> = {
    title: LIST_TITLE,
    description: "",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionPCI(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
    const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
    await ensureListProvision(sp, {
        ...definition,
        lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
    });
}

export default provisionPCI;