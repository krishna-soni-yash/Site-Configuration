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
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.ManagementEffortLog;

type ManagementEffortLogFieldName =
	| "TaskDescription"
	| "ActualStartDate"
	| "ActualEndDate"
	| "ActualEffortHrs"
	| "UpdatedBy"
	| "ManagementTaskID"
	| "Completion"
	| "Remarks";

type ManagementEffortLogViewField = ManagementEffortLogFieldName | "Editor" | "Modified";

function buildFieldDefinitions(taskLogListId: string): FieldDefinition<ManagementEffortLogFieldName>[] {
	return [
		{
			internalName: "TaskDescription",
			schemaXml: `<Field Type='Text' Name='TaskDescription' StaticName='TaskDescription' DisplayName='TaskDescription' MaxLength='255' />`
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
			internalName: "ActualEffortHrs",
			schemaXml: `<Field Type='Number' Name='ActualEffortHrs' StaticName='ActualEffortHrs' DisplayName='ActualEffortHrs' Decimals='2' />`
		},
		{
			internalName: "UpdatedBy",
			schemaXml: `<Field Type='User' Name='UpdatedBy' StaticName='UpdatedBy' DisplayName='UpdatedBy' UserSelectionMode='PeopleOnly' Mult='FALSE' />`
		},
		{
			internalName: "ManagementTaskID",
			schemaXml: `<Field Type='Lookup' Name='ManagementTaskID' StaticName='ManagementTaskID' DisplayName='ManagementTaskID' List='${taskLogListId}' ShowField='ID' LookupId='TRUE' />`
		},
		{
			internalName: "Completion",
			schemaXml: `<Field Type='Boolean' Name='Completion' StaticName='Completion' DisplayName='Completion' Default='0' />`
		},
		{
			internalName: "Remarks",
			schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
		}
	];
}

const defaultViewFields: readonly ManagementEffortLogViewField[] = [
	"TaskDescription",
	"ActualStartDate",
	"ActualEndDate",
	"ActualEffortHrs",
	"UpdatedBy",
	"ManagementTaskID",
	"Completion",
	"Remarks"
] as const;

const removeExistingFields: readonly ManagementEffortLogFieldName[] = [];

export async function provisionManagementEffortLog(sp: SPFI): Promise<void> {
	let managementTaskLogId: string | undefined;

	try {
		const listInfo = await sp.web.lists.getByTitle(RequiredListsProvision.ManagementTaskLog).select("Id")();
		managementTaskLogId = `${listInfo.Id}`;
	} catch (error) {
		const ensureResult = await sp.web.lists.ensure(
			RequiredListsProvision.ManagementTaskLog,
			"Management task log list",
			100
		);
		const ensuredInfo = await ensureResult.list.select("Id")();
		managementTaskLogId = `${ensuredInfo.Id}`;
	}

	if (!managementTaskLogId) {
		throw new Error("Unable to resolve Management Task Log list identifier.");
	}

	if (!managementTaskLogId.startsWith("{")) {
		managementTaskLogId = `{${managementTaskLogId}}`;
	}

	const fields = buildFieldDefinitions(managementTaskLogId);

	const definition: ListProvisionDefinition<ManagementEffortLogFieldName, ManagementEffortLogViewField> = {
		title: LIST_TITLE,
		description: "Management effort log list",
		templateId: 100,
		fields,
		defaultViewFields,
		removeFields: removeExistingFields
	};

	await ensureListProvision(sp, definition);
}

export default provisionManagementEffortLog;
