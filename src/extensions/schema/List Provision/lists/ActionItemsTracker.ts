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
import { provisionMinutesOfMeeting } from "./MinutesOfMeeting";

const LIST_TITLE = RequiredListsProvision.ActionItemsTracker;

type ActionItemsTrackerFieldName =
	| "MoM"
	| "MeetingDate"
	| "ActionItem"
	| "Responsibility"
	| "Plannedclosuredate"
	| "Actualclosuredate"
	| "ClosureDetailsRemarks"
	| "Status";

type ActionItemsTrackerViewField = ActionItemsTrackerFieldName;

function buildFieldDefinitions(minutesListId: string): FieldDefinition<ActionItemsTrackerFieldName>[] {
	return [
		{
			internalName: "MoM",
			schemaXml: `<Field Type='Lookup' Name='MoM' StaticName='MoM' DisplayName='MoM' List='${minutesListId}' ShowField='ID' LookupId='TRUE' />`
		},
		{
			internalName: "MeetingDate",
			schemaXml: `<Field Type='DateTime' Name='MeetingDate' StaticName='MeetingDate' DisplayName='MeetingDate' Format='DateOnly' />`
		},
		{
			internalName: "ActionItem",
			schemaXml: `<Field Type='Text' Name='ActionItem' StaticName='ActionItem' DisplayName='ActionItem' MaxLength='255' />`
		},
		{
			internalName: "Responsibility",
			schemaXml: `<Field Type='User' Name='Responsibility' StaticName='Responsibility' DisplayName='Responsibility' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
		},
		{
			internalName: "Plannedclosuredate",
			schemaXml: `<Field Type='DateTime' Name='Plannedclosuredate' StaticName='Plannedclosuredate' DisplayName='Plannedclosuredate' Format='DateOnly' />`
		},
		{
			internalName: "Actualclosuredate",
			schemaXml: `<Field Type='DateTime' Name='Actualclosuredate' StaticName='Actualclosuredate' DisplayName='Actualclosuredate' Format='DateOnly' />`
		},
		{
			internalName: "ClosureDetailsRemarks",
			schemaXml: `<Field Type='Note' Name='ClosureDetailsRemarks' StaticName='ClosureDetailsRemarks' DisplayName='ClosureDetailsRemarks' NumLines='6' RichText='FALSE' />`
		},
		{
			internalName: "Status",
			schemaXml: `<Field Type='Text' Name='Status' StaticName='Status' DisplayName='Status' MaxLength='255' />`
		}
	];
}

const defaultViewFields: readonly ActionItemsTrackerViewField[] = [
	"MoM",
	"MeetingDate",
	"ActionItem",
	"Responsibility",
	"Plannedclosuredate",
	"Actualclosuredate",
	"ClosureDetailsRemarks",
	"Status"
] as const;

const viewDefinitions: ReadonlyArray<{
	title: string;
	fields: readonly ActionItemsTrackerViewField[];
	makeDefault?: boolean;
	includeLinkTitle?: boolean;
}> = [
	{
		title: "All Action Items",
		fields: defaultViewFields,
		makeDefault: true,
		includeLinkTitle: false
	}
];

async function resolveMinutesOfMeetingListId(sp: SPFI): Promise<string> {
	try {
		const listInfo = await sp.web.lists.getByTitle(RequiredListsProvision.MinutesOfMeeting).select("Id")();
		return normalizeListId(listInfo.Id);
	} catch (error) {
		await provisionMinutesOfMeeting(sp);
		const ensuredInfo = await sp.web.lists.getByTitle(RequiredListsProvision.MinutesOfMeeting).select("Id")();
		return normalizeListId(ensuredInfo.Id);
	}
}

function normalizeListId(id: unknown): string {
	const value = `${id ?? ""}`;
	if (value.length === 0) {
		throw new Error("Unable to resolve Minutes of Meeting list identifier.");
	}
	return value.startsWith("{") ? value : `{${value}}`;
}

export async function provisionActionItemsTracker(sp: SPFI): Promise<void> {
	const minutesListId = await resolveMinutesOfMeetingListId(sp);

	const fields = buildFieldDefinitions(minutesListId);

	const definition: ListProvisionDefinition<ActionItemsTrackerFieldName, ActionItemsTrackerViewField> = {
		title: LIST_TITLE,
		description: "Action items tracker list",
		templateId: 100,
		fields,
		indexedFields: ["Status", "ActionItem"],
		defaultViewFields,
		views: viewDefinitions
	};

	await ensureListProvision(sp, definition);
}

export default provisionActionItemsTracker;
