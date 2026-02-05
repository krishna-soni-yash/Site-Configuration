import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    ensureListProvision,
    ListProvisionDefinition,
    FieldDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.ManagementTaskLog;

type ManagementTaskLogFieldName =
    | "TaskDescription"
    | "Phase"
    | "Activity"
    | "Responsibility"
    | "PlannedStartDate"
    | "PlannedEndDate"
    | "PlannedEffortHrs"
    | "TaskStatus"
    | "Remarks"
    | "Completion"
    | "CompletionCount";

type ManagementTaskLogViewField = ManagementTaskLogFieldName | "Editor" | "Modified";

const fieldDefinitions: readonly FieldDefinition<ManagementTaskLogFieldName>[] = [
    {
        internalName: "Activity",
        schemaXml: `<Field Type='Text' Name='Activity' StaticName='Activity' DisplayName='Activity' MaxLength='255' />`
    },
    {
        internalName: "Responsibility",
        schemaXml: `<Field Type='User' Name='Responsibility' StaticName='Responsibility' DisplayName='Responsibility' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
    },
    {
        internalName: "TaskDescription",
        schemaXml: `<Field Type='Note' Name='TaskDescription' StaticName='TaskDescription' DisplayName='TaskDescription' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "PlannedStartDate",
        schemaXml: `<Field Type='DateTime' Name='PlannedStartDate' StaticName='PlannedStartDate' DisplayName='PlannedStartDate' Format='DateOnly' />`
    },
    {
        internalName: "PlannedEndDate",
        schemaXml: `<Field Type='DateTime' Name='PlannedEndDate' StaticName='PlannedEndDate' DisplayName='PlannedEndDate' Format='DateOnly' />`
    },
    {
        internalName: "PlannedEffortHrs",
        schemaXml: `<Field Type='Number' Name='PlannedEffortHrs' StaticName='PlannedEffortHrs' DisplayName='PlannedEffortHrs' Decimals='2' />`
    },
    {
        internalName: "TaskStatus",
        schemaXml: `<Field Type='Text' Name='TaskStatus' StaticName='TaskStatus' DisplayName='TaskStatus' MaxLength='255' />`
    },
    {
        internalName: "Phase",
        schemaXml: `<Field Type='Text' Name='Phase' StaticName='Phase' DisplayName='Phase' MaxLength='255' />`
    },
    {
        internalName: "Remarks",
        schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "Completion",
        schemaXml: `<Field Type='Boolean' Name='Completion' StaticName='Completion' DisplayName='Completion' Default='0' />`
    },
    {
        internalName: "CompletionCount",
        schemaXml: `<Field Type='Number' Name='CompletionCount' StaticName='CompletionCount' DisplayName='CompletionCount' Decimals='0' Default='0' />`
    }
] as const;

const defaultViewFields: readonly ManagementTaskLogViewField[] = [
    "TaskDescription",
    "Phase",
    "Activity",
    "Responsibility",
    "TaskStatus",
    "PlannedStartDate",
    "PlannedEndDate",
    "PlannedEffortHrs",
    "Remarks",
    "Completion",
    "CompletionCount"
] as const;

const removeExistingFields: readonly ManagementTaskLogFieldName[] = [];

const definition: ListProvisionDefinition<ManagementTaskLogFieldName, ManagementTaskLogViewField> = {
    title: LIST_TITLE,
    description: "Management task log list",
    templateId: 100,
    fields: fieldDefinitions,
    indexedFields: ["TaskStatus", "Completion"],
    defaultViewFields,
    removeFields: removeExistingFields
};

export async function provisionManagementTaskLog(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionManagementTaskLog;
