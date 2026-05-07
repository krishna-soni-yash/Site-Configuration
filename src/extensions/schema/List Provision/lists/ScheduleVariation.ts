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

const LIST_TITLE = RequiredListsProvision.ScheduleVariation;

type ScheduleVariationFieldName =
	| "PlannedStartDate"
	| "PlannedEndDate"
	| "ActualStartDate"
	| "ActualEndDate"
	| "PlannedDuration"
	| "ScheduleVariation"
	| "Goal"
	| "USL"
	| "LSL"
	| "Mean"
	| "Variance"
	| "ProjectType";

type ScheduleVariationViewField = ScheduleVariationFieldName;

const fieldDefinitions: readonly FieldDefinition<ScheduleVariationFieldName>[] = [
	{
		internalName: "PlannedStartDate",
		schemaXml: `<Field Type='DateTime' Name='PlannedStartDate' StaticName='PlannedStartDate' DisplayName='PlannedStartDate' Format='DateOnly' />`
	},
	{
		internalName: "PlannedEndDate",
		schemaXml: `<Field Type='DateTime' Name='PlannedEndDate' StaticName='PlannedEndDate' DisplayName='PlannedEndDate' Format='DateOnly' />`
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
		internalName: "PlannedDuration",
		schemaXml: `<Field Type='Number' Name='PlannedDuration' StaticName='PlannedDuration' DisplayName='PlannedDuration' Decimals='2' />`
	},
	{
		internalName: "ScheduleVariation",
		schemaXml: `<Field Type='Number' Name='ScheduleVariation' StaticName='ScheduleVariation' DisplayName='ScheduleVariation' Decimals='2' />`
	},
	{
		internalName: "Goal",
		schemaXml: `<Field Type='Text' Name='Goal' StaticName='Goal' DisplayName='Goal' MaxLength='255' />`
	},
	{
		internalName: "USL",
		schemaXml: `<Field Type='Number' Name='USL' StaticName='USL' DisplayName='USL' Decimals='2' />`
	},
	{
		internalName: "LSL",
		schemaXml: `<Field Type='Number' Name='LSL' StaticName='LSL' DisplayName='LSL' Decimals='2' />`
	},
	{
		internalName: "Mean",
		schemaXml: `<Field Type='Number' Name='Mean' StaticName='Mean' DisplayName='Mean' Decimals='2' />`
	},
	{
		internalName: "Variance",
		schemaXml: `<Field Type='Number' Name='Variance' StaticName='Variance' DisplayName='Variance' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly ScheduleVariationViewField[] = [
	"PlannedStartDate",
	"PlannedEndDate",
	"ActualStartDate",
	"ActualEndDate",
	"PlannedDuration",
	"ScheduleVariation",
	"Goal",
	"USL",
	"LSL",
	"Mean",
	"Variance",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<ScheduleVariationFieldName, ScheduleVariationViewField> = {
	title: LIST_TITLE,
	description: "Schedule variation",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionScheduleVariation(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionScheduleVariation;
