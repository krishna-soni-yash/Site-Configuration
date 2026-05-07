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

const LIST_TITLE = RequiredListsProvision.OverallProductivity;

type OverallProductivityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "ProjectType"
	| "OverallProductivity"
	| "ActualEffort"
	| "ActualSize"
	| "ActualEfforts_LastMonth"
	| "ActualSize_LastMonth"
	| "OverallProductivity_Mean"
	| "Variance";

type OverallProductivityViewField = OverallProductivityFieldName;

const fieldDefinitions: readonly FieldDefinition<OverallProductivityFieldName>[] = [
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
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	},
	{
		internalName: "OverallProductivity",
		schemaXml: `<Field Type='Number' Name='OverallProductivity' StaticName='OverallProductivity' DisplayName='OverallProductivity' Decimals='2' />`
	},
	{
		internalName: "ActualEffort",
		schemaXml: `<Field Type='Number' Name='ActualEffort' StaticName='ActualEffort' DisplayName='ActualEffort' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "ActualEfforts_LastMonth",
		schemaXml: `<Field Type='Number' Name='ActualEfforts_LastMonth' StaticName='ActualEfforts_LastMonth' DisplayName='ActualEfforts_LastMonth' Decimals='2' />`
	},
	{
		internalName: "ActualSize_LastMonth",
		schemaXml: `<Field Type='Number' Name='ActualSize_LastMonth' StaticName='ActualSize_LastMonth' DisplayName='ActualSize_LastMonth' Decimals='2' />`
	},
	{
		internalName: "OverallProductivity_Mean",
		schemaXml: `<Field Type='Number' Name='OverallProductivity_Mean' StaticName='OverallProductivity_Mean' DisplayName='OverallProductivity_Mean' Decimals='2' />`
	},
	{
		internalName: "Variance",
		schemaXml: `<Field Type='Number' Name='Variance' StaticName='Variance' DisplayName='Variance' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly OverallProductivityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"ProjectType",
	"OverallProductivity",
	"ActualEffort",
	"ActualSize",
	"ActualEfforts_LastMonth",
	"ActualSize_LastMonth",
	"OverallProductivity_Mean",
	"Variance"
] as const;

const definition: ListProvisionDefinition<OverallProductivityFieldName, OverallProductivityViewField> = {
	title: LIST_TITLE,
	description: "Overall productivity",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionOverallProductivity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionOverallProductivity;
