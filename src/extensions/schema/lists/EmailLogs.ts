import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    RequiredListsProvision,
    ensureListProvision,
    ListProvisionDefinition,
    FieldDefinition
} from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.EmailLogs;

type EmailLogsFieldName = "Status" | "MailSentTo" | "VersionId" | "MailSentStatus" | "TestDesc";
type EmailLogsViewField = EmailLogsFieldName | "LinkTitle" | "Author" | "Created";

const fieldDefinitions: readonly FieldDefinition<EmailLogsFieldName>[] = [
    {
        internalName: "Status",
        schemaXml: `<Field Type='Text' Name='Status' StaticName='Status' DisplayName='Status' MaxLength='255' />`
    },
    {
        internalName: "MailSentTo",
        schemaXml: `<Field Type='Text' Name='MailSentTo' StaticName='MailSentTo' DisplayName='MailSentTo' MaxLength='255' />`
    },
    {
        internalName: "VersionId",
        schemaXml: `<Field Type='Text' Name='VersionId' StaticName='VersionId' DisplayName='VersionId' MaxLength='255' />`
    },
    {
        internalName: "MailSentStatus",
        schemaXml: `<Field Type='Note' Name='MailSentStatus' StaticName='MailSentStatus' DisplayName='MailSentStatus' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "TestDesc",
        schemaXml: `<Field Type='Text' Name='TestDesc' StaticName='TestDesc' DisplayName='TestDesc' MaxLength='255' />`
    }
] as const;

const defaultViewFields: readonly EmailLogsViewField[] = [
    "LinkTitle",
    "Status",
    "MailSentTo",
    "VersionId",
    "MailSentStatus",
    "Author",
    "Created"
] as const;

const removeExistingFields: readonly EmailLogsFieldName[] = [];

const definition: ListProvisionDefinition<EmailLogsFieldName, EmailLogsViewField> = {
    title: LIST_TITLE,
    description: "Email logs list",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields,
    removeFields: removeExistingFields
};

export async function provisionEmailLogs(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionEmailLogs;