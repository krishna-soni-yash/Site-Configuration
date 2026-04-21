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

const LIST_TITLE = RequiredListsProvision.CostOfQuality;

type CostOfQualityFieldName =
	| "PreventionCost"
	| "AppraisalCost"
	| "FailureCost"
	| "COQ"
	| "TotalProjectEffort"
	| "COQPercentage"
	| "Goal"
	| "USL"
	| "LSL"
	| "ProjectType";

type CostOfQualityViewField = CostOfQualityFieldName;

const fieldDefinitions: readonly FieldDefinition<CostOfQualityFieldName>[] = [
	{
		internalName: "PreventionCost",
		schemaXml: `<Field Type='Number' Name='PreventionCost' StaticName='PreventionCost' DisplayName='PreventionCost' Decimals='2' />`
	},
	{
		internalName: "AppraisalCost",
		schemaXml: `<Field Type='Number' Name='AppraisalCost' StaticName='AppraisalCost' DisplayName='AppraisalCost' Decimals='2' />`
	},
	{
		internalName: "FailureCost",
		schemaXml: `<Field Type='Number' Name='FailureCost' StaticName='FailureCost' DisplayName='FailureCost' Decimals='2' />`
	},
	{
		internalName: "COQ",
		schemaXml: `<Field Type='Number' Name='COQ' StaticName='COQ' DisplayName='COQ' Decimals='2' />`
	},
	{
		internalName: "TotalProjectEffort",
		schemaXml: `<Field Type='Number' Name='TotalProjectEffort' StaticName='TotalProjectEffort' DisplayName='TotalProjectEffort' Decimals='2' />`
	},
	{
		internalName: "COQPercentage",
		schemaXml: `<Field Type='Number' Name='COQPercentage' StaticName='COQPercentage' DisplayName='COQPercentage' Decimals='2' />`
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
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly CostOfQualityViewField[] = [
	"PreventionCost",
	"AppraisalCost",
	"FailureCost",
	"COQ",
	"TotalProjectEffort",
	"COQPercentage",
	"Goal",
	"USL",
	"LSL",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<CostOfQualityFieldName, CostOfQualityViewField> = {
	title: LIST_TITLE,
	description: "Cost of Quality",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionCostOfQuality(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionCostOfQuality;
