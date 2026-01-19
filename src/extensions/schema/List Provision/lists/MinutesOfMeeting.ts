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

const LIST_TITLE = RequiredListsProvision.MinutesOfMeeting;

type MinutesOfMeetingFieldName =
	| "MeetingName"
	| "Venue"
	| "Agenda"
	| "Participants"
	| "Summary"
	| "MeetingDate"
	| "CalledBy"
	| "StartTime"
	| "EndTime"
	| "MinutedBy";

type MinutesOfMeetingViewField = MinutesOfMeetingFieldName;

const fieldDefinitions: readonly FieldDefinition<MinutesOfMeetingFieldName>[] = [
	{
		internalName: "MeetingName",
		schemaXml: `<Field Type='Text' Name='MeetingName' StaticName='MeetingName' DisplayName='MeetingName' MaxLength='255' />`
	},
	{
		internalName: "Venue",
		schemaXml: `<Field Type='Text' Name='Venue' StaticName='Venue' DisplayName='Venue' MaxLength='255' />`
	},
	{
		internalName: "Agenda",
		schemaXml: `<Field Type='Text' Name='Agenda' StaticName='Agenda' DisplayName='Agenda' MaxLength='255' />`
	},
	{
		internalName: "Participants",
		schemaXml: `<Field Type='User' Name='Participants' StaticName='Participants' DisplayName='Participants' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	},
	{
		internalName: "Summary",
		schemaXml: `<Field Type='Note' Name='Summary' StaticName='Summary' DisplayName='Summary' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "MeetingDate",
		schemaXml: `<Field Type='DateTime' Name='MeetingDate' StaticName='MeetingDate' DisplayName='MeetingDate' Format='DateOnly' />`
	},
	{
		internalName: "CalledBy",
		schemaXml: `<Field Type='User' Name='CalledBy' StaticName='CalledBy' DisplayName='CalledBy' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	},
	{
		internalName: "StartTime",
		schemaXml: `<Field Type='DateTime' Name='StartTime' StaticName='StartTime' DisplayName='StartTime' Format='DateTime' />`
	},
	{
		internalName: "EndTime",
		schemaXml: `<Field Type='DateTime' Name='EndTime' StaticName='EndTime' DisplayName='EndTime' Format='DateTime' />`
	},
	{
		internalName: "MinutedBy",
		schemaXml: `<Field Type='User' Name='MinutedBy' StaticName='MinutedBy' DisplayName='MinutedBy' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	}
] as const;

const defaultViewFields: readonly MinutesOfMeetingViewField[] = [
	"MeetingName",
	"Venue",
	"Agenda",
	"Participants",
	"Summary",
	"MeetingDate",
	"CalledBy",
	"StartTime",
	"EndTime",
	"MinutedBy"
] as const;

const viewDefinitions: ReadonlyArray<{
	title: string;
	fields: readonly MinutesOfMeetingViewField[];
	makeDefault?: boolean;
	includeLinkTitle?: boolean;
}> = [
	{
		title: "All Minutes",
		fields: defaultViewFields,
		makeDefault: true,
		includeLinkTitle: false
	}
];

const definition: ListProvisionDefinition<MinutesOfMeetingFieldName, MinutesOfMeetingViewField> = {
	title: LIST_TITLE,
	description: "Minutes of meeting list",
	templateId: 100,
	fields: fieldDefinitions,
	indexedFields: ["MeetingName"],
	defaultViewFields,
	views: viewDefinitions
};

export async function provisionMinutesOfMeeting(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionMinutesOfMeeting;
