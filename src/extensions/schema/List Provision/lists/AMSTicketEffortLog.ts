/*eslint-disable*/
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
import { provisionAMSTicketLog } from "./AMSTicketLog";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.AMSTicketEffortLog;

type AMSTicketEffortLogFieldName =
    | "TicketID"
    | "TicketID_Title"
    | "TicketID_TicketDescription"
    | "AssignedTo"
    | "UniqueEffortID"
    | "TaskType"
    | "TaskStatus"
    | "ActualEffort"
    | "ActualStartDate"
    | "ActualEndDate"
    | "Remarks";

type AMSTicketEffortLogViewField = AMSTicketEffortLogFieldName | "ID" | "Modified" | "Editor";

const taskTypeChoices = ["Solution Analysis Effort", "Solution Effort", "Solution Review Effort", "Solution Rework Effort", "Solution Testing Effort", "Others"] as const;
const taskStatusChoices = ["Open", "Closed", "On-hold", "In progress"] as const;

function buildChoiceFieldSchema(
    name: string,
    displayName: string,
    choices: readonly string[],
    options?: { multi?: boolean }
): string {
    const type = options?.multi ? "MultiChoice" : "Choice";
    const format = options?.multi ? "Checkboxes" : "Dropdown";
    const choicesXml = choices.map((c) => `<CHOICE>${c}</CHOICE>`).join("");
    return `<Field Type='${type}' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='${format}'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

function normalizeListId(id: unknown): string {
    const value = `${id ?? ""}`;
    if (value.length === 0) {
        throw new Error("Unable to resolve AMSTicketLog list identifier.");
    }
    return value.startsWith("{") ? value : `{${value}}`;
}

async function resolveAMSTicketListId(sp: SPFI): Promise<string> {
    try {
        const info = await sp.web.lists.getByTitle("AMSTicketLog").select("Id")();
        return normalizeListId(info.Id);
    } catch (err) {
        await provisionAMSTicketLog(sp);
        const ensured = await sp.web.lists.getByTitle("AMSTicketLog").select("Id")();
        return normalizeListId(ensured.Id);
    }
}

function buildFieldDefinitions(ticketListId: string): FieldDefinition<AMSTicketEffortLogFieldName>[] {
    return [
        {
            internalName: "TicketID",
            schemaXml: `<Field Type='Lookup' Name='TicketID' StaticName='TicketID' DisplayName='TicketID' List='${ticketListId}' ShowField='ID' LookupId='TRUE' />`
        },
        {
            internalName: "TicketID_Title",
            schemaXml: `<Field Type='Lookup' Name='TicketID_Title' StaticName='TicketID_Title' DisplayName='TicketTitle' List='${ticketListId}' ShowField='Title' />`
        },
        {
            internalName: "TicketID_TicketDescription",
            schemaXml: `<Field Type='Lookup' Name='TicketID_TicketDescription' StaticName='TicketID_TicketDescription' DisplayName='TicketDescription' List='${ticketListId}' ShowField='TicketDescription' />`
        },
        {
            internalName: "AssignedTo",
            schemaXml: `<Field Type='User' Name='AssignedTo' StaticName='AssignedTo' DisplayName='AssignedTo' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
        },
        {
            internalName: "UniqueEffortID",
            schemaXml: `<Field Type='Text' Name='UniqueEffortID' StaticName='UniqueEffortID' DisplayName='UniqueEffortID' MaxLength='255' />`
        },
        {
            internalName: "TaskType",
            schemaXml: buildChoiceFieldSchema("TaskType", "TaskType", taskTypeChoices as readonly string[], { multi: true })
        },
        {
            internalName: "TaskStatus",
            schemaXml: buildChoiceFieldSchema("TaskStatus", "TaskStatus", taskStatusChoices as readonly string[])
        },
        {
            internalName: "ActualEffort",
            schemaXml: `<Field Type='Number' Name='ActualEffort' StaticName='ActualEffort' DisplayName='ActualEffort' Decimals='2' />`
        },
        {
            internalName: "ActualStartDate",
            schemaXml: `<Field Type='DateTime' Name='ActualStartDate' StaticName='ActualStartDate' DisplayName='ActualStartDate' Format='DateOnly' />`
        },
        {
            internalName: "ActualEndDate",
            schemaXml: `<Field Type='DateTime' Name='ActualEndDate' StaticName='ActualEndDate' DisplayName='ActualEndDate' Format='DateOnly' />`
        },
        {
            internalName: "Remarks",
            schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
        }
    ];
}

const defaultViewFields: readonly AMSTicketEffortLogViewField[] = [
    "TicketID",
    "TicketID_Title",
    "TicketID_TicketDescription",
    "AssignedTo",
    "UniqueEffortID",
    "TaskType",
    "TaskStatus",
    "ActualEffort",
    "ActualStartDate",
    "ActualEndDate",
    "Remarks"
] as const;

export async function provisionAMSTicketEffortLog(sp: SPFI): Promise<void> {
    const ticketListId = await resolveAMSTicketListId(sp);
    const fields = buildFieldDefinitions(ticketListId);

    const definition: ListProvisionDefinition<AMSTicketEffortLogFieldName, AMSTicketEffortLogViewField> = {
        title: LIST_TITLE,
        description: "AMS ticket effort log",
        templateId: 100,
        fields,
        defaultViewFields
    };

    await ensureListProvision(sp, definition);
}

export default provisionAMSTicketEffortLog;
