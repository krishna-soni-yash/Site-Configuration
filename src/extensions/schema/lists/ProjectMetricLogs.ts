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

const LIST_TITLE = RequiredListsProvision.ProjectMetricLogs;

type ProjectMetricLogsFieldName =
    | "VersionId"
    | "Status"
    | "PMComments"
    | "ReviewerComments"
    | "CreatedVersion"
    | "IsActive";

const statusChoices = ["Draft", "In Review", "In Approval", "Approved", "Rejected"] as const;
const createdVersionChoices = ["Minor", "Major"] as const;

const fieldDefinitions: readonly FieldDefinition<ProjectMetricLogsFieldName>[] = [
    {
        internalName: "VersionId",
        schemaXml: `<Field Type='Text' Name='VersionId' StaticName='VersionId' DisplayName='VersionId' MaxLength='255' />`
    },
    {
        internalName: "Status",
        schemaXml: `<Field Type='Choice' Name='Status' StaticName='Status' DisplayName='Status' Format='Dropdown'><CHOICES>${statusChoices
            .map(choice => `<CHOICE>${choice}</CHOICE>`)
            .join("")}</CHOICES></Field>`
    },
    {
        internalName: "PMComments",
        schemaXml: `<Field Type='Note' Name='PMComments' StaticName='PMComments' DisplayName='PMComments' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "ReviewerComments",
        schemaXml: `<Field Type='Note' Name='ReviewerComments' StaticName='ReviewerComments' DisplayName='ReviewerComments' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "CreatedVersion",
        schemaXml: `<Field Type='Choice' Name='CreatedVersion' StaticName='CreatedVersion' DisplayName='CreatedVersion' Format='Dropdown'><CHOICES>${createdVersionChoices
            .map(choice => `<CHOICE>${choice}</CHOICE>`)
            .join("")}</CHOICES></Field>`
    },
    {
        internalName: "IsActive",
        schemaXml: `<Field Type='Boolean' Name='IsActive' StaticName='IsActive' DisplayName='IsActive' />`
    }
] as const;

const defaultViewFields: readonly ProjectMetricLogsFieldName[] = [
    "VersionId",
    "Status",
    "PMComments",
    "ReviewerComments",
    "CreatedVersion",
    "IsActive"
] as const;

const definition: ListProvisionDefinition<ProjectMetricLogsFieldName> = {
    title: LIST_TITLE,
    description: "Project metrics logs list",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionProjectMetricLogs(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionProjectMetricLogs;