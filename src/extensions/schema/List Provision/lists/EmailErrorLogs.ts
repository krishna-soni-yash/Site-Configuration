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

const LIST_TITLE = RequiredListsProvision.EmailErrorLogs;

type EmailErrorLogsFieldName = "ErrorDescription" | "FlowName";
type EmailErrorLogsViewField = EmailErrorLogsFieldName | "ID" | "Modified" | "Editor";

const fieldDefinitions: readonly FieldDefinition<EmailErrorLogsFieldName>[] = [
    {
        internalName: "ErrorDescription",
        schemaXml: `<Field Type='Note' Name='ErrorDescription' StaticName='ErrorDescription' DisplayName='ErrorDescription' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "FlowName",
        schemaXml: `<Field Type='Text' Name='FlowName' StaticName='FlowName' DisplayName='FlowName' MaxLength='255' />`
    }
] as const;

const defaultViewFields: readonly EmailErrorLogsViewField[] = [
    "ID",
    "ErrorDescription",
    "FlowName",
    "Modified",
    "Editor"
] as const;

const definition: ListProvisionDefinition<EmailErrorLogsFieldName, EmailErrorLogsViewField> = {
    title: LIST_TITLE,
    description: "Email error logs",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionEmailErrorLogs(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionEmailErrorLogs;
