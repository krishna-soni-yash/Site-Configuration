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
import { provisionWorkLogManagement } from "./WorkLogManagement";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.TaskManagement;

type TaskManagementFieldName =
    | "WorkItemNo"
    | "WorkItemNo_ReqTitle"
    | "WorkItemNo_WorkItemNo"
    | "UniqueTaskID"
    | "TaskType"
    | "AssignedTo"
    | "ActualEffort"
    | "ActualEndDate"
    | "ActualStartDate"
    | "Remarks"
    | "TaskStatus";

type TaskManagementViewField = TaskManagementFieldName | "ID" | "Modified" | "Editor";

const taskTypeChoices = [
    "Requirements Analysis",
    "Design",
    "Coding",
    "Unit Testing",
    "Rework after Unit Testing",
    "Code Review",
    "Rework after Code Review",
    "System Testing",
    "Rework after System Testing",
    "Integration Testing",
    "Rework after Integration Testing",
    "Re-testing",
    "Acceptance testing",
    "Preparation of Test Cases",
    "Review of Test Cases",
    "Rework on Test Cases",
    "Baseline Test Cases",
    "Creation of SRS",
    "Review of SRS",
    "Rework on SRS",
    "Baseline SRS",
    "Preparation of Design Document",
    "Review of Design Document",
    "Rework on Design Document",
    "Baseline Design Document",
    "Preparation of Unit Test Cases",
    "Review of Unit Test Cases",
    "Rework on Unit Test Cases",
    "Baseline Unit Test Cases",
    "Create Release Notes",
    "Review Release Notes",
    "Rework Release Notes",
    "Baseline Release Notes",
    "Create Build",
    "Review Build",
    "Rework Build",
    "Baseline Build",
    "Smoke Test Build",
    "Preparation of Test Plan",
    "Review of Test Plan",
    "Rework on Test Plan",
    "Baseline Test Plan",
    "Test Environment Setup",
    "Review of Test Environment Setup",
    "Creation of Traceability Matrix",
    "Updation of Traceability Matrix",
    "Create Delivery Notes",
    "Review  Delivery Notes",
    "Rework  Delivery Notes",
    "Baseline  Delivery Notes",
    "Delivery Inspection"
] as const;

const taskStatusChoices = ["Not Started", "In Progress", "Completed", "Deferred", "Waiting on someone else", "Open"] as const;

function buildChoiceFieldSchema(name: string, displayName: string, choices: readonly string[]): string {
    const choicesXml = choices.map((c) => `<CHOICE>${c}</CHOICE>`).join("");
    return `<Field Type='Choice' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='Dropdown'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

function normalizeListId(id: unknown): string {
    const value = `${id ?? ""}`;
    if (value.length === 0) {
        throw new Error("Unable to resolve WorkLogManagement list identifier.");
    }
    return value.startsWith("{") ? value : `{${value}}`;
}

async function resolveWorkLogListId(sp: SPFI): Promise<string> {
    try {
        const info = await sp.web.lists.getByTitle("WorkLogManagement").select("Id")();
        return normalizeListId(info.Id);
    } catch (err) {
        await provisionWorkLogManagement(sp);
        const ensured = await sp.web.lists.getByTitle("WorkLogManagement").select("Id")();
        return normalizeListId(ensured.Id);
    }
}

function buildFieldDefinitions(workLogListId: string): FieldDefinition<TaskManagementFieldName>[] {
    return [
        {
            internalName: "WorkItemNo",
            schemaXml: `<Field Type='Lookup' Name='WorkItemNo' StaticName='WorkItemNo' DisplayName='WorkItemNo' List='${workLogListId}' ShowField='ID' LookupId='TRUE' />`
        },
        {
            internalName: "WorkItemNo_ReqTitle",
            schemaXml: `<Field Type='Lookup' Name='WorkItemNo_ReqTitle' StaticName='WorkItemNo_ReqTitle' DisplayName='WorkItemReqTitle' List='${workLogListId}' ShowField='ReqTitle' />`
        },
        {
            internalName: "WorkItemNo_WorkItemNo",
            schemaXml: `<Field Type='Lookup' Name='WorkItemNo_WorkItemNo' StaticName='WorkItemNo_WorkItemNo' DisplayName='WorkItemWorkItemNo' List='${workLogListId}' ShowField='WorkItemNo' />`
        },
        {
            internalName: "UniqueTaskID",
            schemaXml: `<Field Type='Text' Name='UniqueTaskID' StaticName='UniqueTaskID' DisplayName='UniqueTaskID' MaxLength='255' />`
        },
        {
            internalName: "TaskType",
            schemaXml: buildChoiceFieldSchema("TaskType", "TaskType", taskTypeChoices as readonly string[])
        }
        ,
        {
            internalName: "AssignedTo",
            schemaXml: `<Field Type='User' Name='AssignedTo' StaticName='AssignedTo' DisplayName='AssignedTo' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
        },
        {
            internalName: "ActualEffort",
            schemaXml: `<Field Type='Number' Name='ActualEffort' StaticName='ActualEffort' DisplayName='ActualEffort' Decimals='2' />`
        },
        {
            internalName: "ActualEndDate",
            schemaXml: `<Field Type='DateTime' Name='ActualEndDate' StaticName='ActualEndDate' DisplayName='ActualEndDate' Format='DateOnly' />`
        },
        {
            internalName: "ActualStartDate",
            schemaXml: `<Field Type='DateTime' Name='ActualStartDate' StaticName='ActualStartDate' DisplayName='ActualStartDate' Format='DateOnly' />`
        },
        {
            internalName: "Remarks",
            schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
        },
        {
            internalName: "TaskStatus",
            schemaXml: buildChoiceFieldSchema("TaskStatus", "TaskStatus", taskStatusChoices as readonly string[])
        }
    ];
}

const defaultViewFields: readonly TaskManagementViewField[] = [
    "WorkItemNo",
    "WorkItemNo_ReqTitle",
    "WorkItemNo_WorkItemNo",
    "UniqueTaskID",
    "TaskType",
    "AssignedTo",
    "ActualEffort",
    "ActualStartDate",
    "ActualEndDate",
    "Remarks",
    "TaskStatus"
] as const;

export async function provisionTaskManagement(sp: SPFI): Promise<void> {
    const workLogListId = await resolveWorkLogListId(sp);
    const fields = buildFieldDefinitions(workLogListId);

    const definition: ListProvisionDefinition<TaskManagementFieldName, TaskManagementViewField> = {
        title: LIST_TITLE,
        description: "Task management list",
        templateId: 100,
        fields,
        defaultViewFields
    };

    await ensureListProvision(sp, definition);
}

export default provisionTaskManagement;
