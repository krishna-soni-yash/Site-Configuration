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

type CustomerSatisfactionFieldName = "Remarks" | "CSATAquiredDate";
type CustomerSatisfactionViewField = CustomerSatisfactionFieldName | "ID" | "Modified" | "Editor" | "Title";

const fieldDefinitions: readonly FieldDefinition<CustomerSatisfactionFieldName>[] = [
    {
        internalName: "Remarks",
        schemaXml: `<Field Type='Text' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' MaxLength='255' />`
    },
    {
        internalName: "CSATAquiredDate",
        schemaXml: `<Field Type='DateTime' Name='CSATAquiredDate' StaticName='CSATAquiredDate' DisplayName='CSATAquiredDate' Format='DateOnly' />`
    }
] as const;

const defaultViewFields: readonly CustomerSatisfactionViewField[] = [
    "ID",
    "Title",
    "Remarks",
    "CSATAquiredDate",
    "Modified",
    "Editor"
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
