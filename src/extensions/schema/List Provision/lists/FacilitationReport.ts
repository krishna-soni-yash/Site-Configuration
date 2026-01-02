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

const LIST_TITLE = RequiredListsProvision.FacilitationReport;

type FacilitationReportFieldName =
	| "Category"
	| "Finding"
	| "FindingDate"
	| "Assigned"
	| "CorrectiveAction"
	| "ClosureDate"
	| "Status"
	| "Remarks"
	| "ManagementTaskID"
    | "Attachments";

type FacilitationReportViewField = FacilitationReportFieldName;

function buildFieldDefinitions(managementTaskLogListId: string): FieldDefinition<FacilitationReportFieldName>[] {
	return [
		{
			internalName: "Category",
			schemaXml: `<Field Type='Text' Name='Category' StaticName='Category' DisplayName='Category' MaxLength='255' />`
		},
		{
			internalName: "Finding",
			schemaXml: `<Field Type='Text' Name='Finding' StaticName='Finding' DisplayName='Finding' MaxLength='255' />`
		},
		{
			internalName: "FindingDate",
			schemaXml: `<Field Type='DateTime' Name='FindingDate' StaticName='FindingDate' DisplayName='FindingDate' Format='DateOnly' />`
		},
		{
			internalName: "Assigned",
			schemaXml: `<Field Type='User' Name='Assigned' StaticName='Assigned' DisplayName='Assigned' UserSelectionMode='PeopleOnly' Mult='FALSE' />`
		},
		{
			internalName: "CorrectiveAction",
			schemaXml: `<Field Type='Note' Name='CorrectiveAction' StaticName='CorrectiveAction' DisplayName='CorrectiveAction' NumLines='6' RichText='FALSE' />`
		},
		{
			internalName: "ClosureDate",
			schemaXml: `<Field Type='DateTime' Name='ClosureDate' StaticName='ClosureDate' DisplayName='ClosureDate' Format='DateOnly' />`
		},
		{
			internalName: "Status",
			schemaXml: `<Field Type='Text' Name='Status' StaticName='Status' DisplayName='Status' MaxLength='255' />`
		},
		{
			internalName: "Remarks",
			schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
		},
		{
			internalName: "ManagementTaskID",
			schemaXml: `<Field Type='Lookup' Name='ManagementTaskID' StaticName='ManagementTaskID' DisplayName='ManagementTaskID' List='${managementTaskLogListId}' ShowField='ID' LookupId='TRUE' />`
		}
	];
}

const defaultViewFields: readonly FacilitationReportViewField[] = [
	"Category",
	"Finding",
	"FindingDate",
	"Assigned",
	"CorrectiveAction",
	"ClosureDate",
	"Status",
	"Remarks",
	"ManagementTaskID",
    "Attachments"
] as const;

const removeExistingFields: readonly FacilitationReportFieldName[] = [];

async function ensureManagementTaskLogListId(sp: SPFI): Promise<string> {
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

	return managementTaskLogId;
}

async function setDefaultViewFieldsOnly(
	sp: SPFI,
	listTitle: string,
	fields: readonly FacilitationReportViewField[]
): Promise<void> {
	const list = sp.web.lists.getByTitle(listTitle);
	const defaultView = list.defaultView;

	try {
		await defaultView.fields.removeAll();
	} catch (error) {
		console.warn(`Failed to clear default view fields for list ${listTitle}:`, error);
	}

	for (const field of fields) {
		try {
			await defaultView.fields.add(field as any);
		} catch (error) {
			console.warn(`Failed to add field ${field} to default view for list ${listTitle}:`, error);
		}
	}
}

export async function provisionFacilitationReport(sp: SPFI): Promise<void> {
	const managementTaskLogId = await ensureManagementTaskLogListId(sp);
	const fields = buildFieldDefinitions(managementTaskLogId);

	const definition: ListProvisionDefinition<FacilitationReportFieldName, FacilitationReportViewField> = {
		title: LIST_TITLE,
		description: "Facilitation report list",
		templateId: 100,
		fields,
		defaultViewFields,
		removeFields: removeExistingFields
	};

	await ensureListProvision(sp, definition);
	await setDefaultViewFieldsOnly(sp, LIST_TITLE, defaultViewFields);
}

export default provisionFacilitationReport;
