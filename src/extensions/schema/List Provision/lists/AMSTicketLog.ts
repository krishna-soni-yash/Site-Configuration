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

const LIST_TITLE = RequiredListsProvision.AMSTicketLog;

type AMSTicketLogFieldName =
    | "TicketDescription"
    | "ReceivedDate"
    | "Priority"
    | "Status"
    | "ClosureDate"
    | "Remarks"
    | "AssignedToUsers";

type AMSTicketLogViewField = AMSTicketLogFieldName | "ID" | "Title" | "Modified" | "Editor";

const priorityChoices = ["P1", "P2", "P3", "P4"] as const;
const statusChoices = ["Open", "Closed", "Resolved", "Hold"] as const;

function buildChoiceFieldSchema(name: string, displayName: string, choices: readonly string[]): string {
    const choicesXml = choices.map((c) => `<CHOICE>${c}</CHOICE>`).join("");
    return `<Field Type='Choice' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='Dropdown'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

const fieldDefinitions: readonly FieldDefinition<AMSTicketLogFieldName>[] = [
    {
        internalName: "TicketDescription",
        schemaXml: `<Field Type='Text' Name='TicketDescription' StaticName='TicketDescription' DisplayName='TicketDescription' MaxLength='255' />`
    },
    {
        internalName: "ReceivedDate",
        schemaXml: `<Field Type='DateTime' Name='ReceivedDate' StaticName='ReceivedDate' DisplayName='ReceivedDate' Format='DateOnly' />`
    },
    {
        internalName: "Priority",
        schemaXml: buildChoiceFieldSchema("Priority", "Priority", priorityChoices as readonly string[])
    },
    {
        internalName: "Status",
        schemaXml: buildChoiceFieldSchema("Status", "Status", statusChoices as readonly string[])
    },
    {
        internalName: "ClosureDate",
        schemaXml: `<Field Type='DateTime' Name='ClosureDate' StaticName='ClosureDate' DisplayName='ClosureDate' Format='DateOnly' />`
    },
    {
        internalName: "Remarks",
        schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "AssignedToUsers",
        schemaXml: `<Field Type='User' Name='AssignedToUsers' StaticName='AssignedToUsers' DisplayName='AssignedToUsers' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
    }
] as const;

const defaultViewFields: readonly AMSTicketLogViewField[] = [
    "ID",
    "Title",
    "TicketDescription",
    "ReceivedDate",
    "Priority",
    "Status",
    "ClosureDate",
    "Remarks",
    "AssignedToUsers",
    "Modified",
    "Editor"
] as const;

const definition: ListProvisionDefinition<AMSTicketLogFieldName, AMSTicketLogViewField> = {
    title: LIST_TITLE,
    description: "AMS ticket log",
    templateId: 100,
    fields: fieldDefinitions,
    indexedFields: ["TicketID","Status"],
    defaultViewFields
};

export async function provisionAMSTicketLog(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
    await sp.web.lists.getByTitle(LIST_TITLE).fields.getByInternalNameOrTitle("Title").update({ Title: "TicketID" });
}

export default provisionAMSTicketLog;
