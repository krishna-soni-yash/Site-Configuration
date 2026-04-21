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

const LIST_TITLE = RequiredListsProvision.CustomerSatisfactionIndex;

type CustomerSatisfactionFieldName = "Remarks" | "Goal" | "USL" | "LSL" | "CSATAquiredDate";
type CustomerSatisfactionViewField = CustomerSatisfactionFieldName |"Title";

const fieldDefinitions: readonly FieldDefinition<CustomerSatisfactionFieldName>[] = [
    {
        internalName: "Remarks",
        schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "Goal",
        schemaXml: `<Field Type='Text' Name='Goal' StaticName='Goal' DisplayName='Goal' />`
    },
    {
        internalName: "USL",
        schemaXml: `<Field Type='Number' Name='USL' StaticName='USL' DisplayName='USL' />`
    },
    {
        internalName: "LSL",
        schemaXml: `<Field Type='Number' Name='LSL' StaticName='LSL' DisplayName='LSL' />`
    },
    {
        internalName: "CSATAquiredDate",
        schemaXml: `<Field Type='DateTime' Name='CSATAquiredDate' StaticName='CSATAquiredDate' DisplayName='CSATAquiredDate' Format='DateOnly' />`
    }
] as const;

const defaultViewFields: readonly CustomerSatisfactionViewField[] = [
    "Remarks",
    "Goal",
    "USL",
    "LSL",
    "CSATAquiredDate"
] as const;

const definition: ListProvisionDefinition<CustomerSatisfactionFieldName, CustomerSatisfactionViewField> = {
    title: LIST_TITLE,
    description: "Customer Satisfaction Index",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionCustomerSatisfactionIndex(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
    await sp.web.lists.getByTitle(LIST_TITLE).fields.getByInternalNameOrTitle("Title").update({ Title: "CSAT Value" });
}

export default provisionCustomerSatisfactionIndex;
