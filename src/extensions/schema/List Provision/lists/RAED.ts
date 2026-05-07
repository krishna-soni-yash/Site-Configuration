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

const LIST_TITLE = RequiredListsProvision.RAED;

type RAEDFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "ProjectType"
	| "RequirementsAnalysisEffortDensit"
	| "ActualSize"
	| "ActualRequirementsAnalysisEffort"
	| "RAED_MR"
	| "RAED_Mean"
	| "RAED_LCL"
	| "RAED_UCL"
	| "RAED_MR_LCL"
	| "RAED_MR_Mean"
	| "RAED_MR_UCL";

type RAEDViewField = RAEDFieldName;

const fieldDefinitions: readonly FieldDefinition<RAEDFieldName>[] = [
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
		internalName: "RequirementsAnalysisEffortDensit",
		schemaXml: `<Field Type='Number' Name='RequirementsAnalysisEffortDensit' StaticName='RequirementsAnalysisEffortDensit' DisplayName='RequirementsAnalysisEffortDensity' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "ActualRequirementsAnalysisEffort",
		schemaXml: `<Field Type='Number' Name='ActualRequirementsAnalysisEffort' StaticName='ActualRequirementsAnalysisEffort' DisplayName='ActualRequirementsAnalysisEffort' Decimals='2' />`
	},
	{
		internalName: "RAED_MR",
		schemaXml: `<Field Type='Number' Name='RAED_MR' StaticName='RAED_MR' DisplayName='RAED_MR' Decimals='2' />`
	},
	{
		internalName: "RAED_Mean",
		schemaXml: `<Field Type='Number' Name='RAED_Mean' StaticName='RAED_Mean' DisplayName='RAED_Mean' Decimals='2' />`
	},
	{
		internalName: "RAED_LCL",
		schemaXml: `<Field Type='Number' Name='RAED_LCL' StaticName='RAED_LCL' DisplayName='RAED_LCL' Decimals='2' />`
	},
	{
		internalName: "RAED_UCL",
		schemaXml: `<Field Type='Number' Name='RAED_UCL' StaticName='RAED_UCL' DisplayName='RAED_UCL' Decimals='2' />`
	},
	{
		internalName: "RAED_MR_LCL",
		schemaXml: `<Field Type='Number' Name='RAED_MR_LCL' StaticName='RAED_MR_LCL' DisplayName='RAED_MR_LCL' Decimals='2' />`
	},
	{
		internalName: "RAED_MR_Mean",
		schemaXml: `<Field Type='Number' Name='RAED_MR_Mean' StaticName='RAED_MR_Mean' DisplayName='RAED_MR_Mean' Decimals='2' />`
	},
	{
		internalName: "RAED_MR_UCL",
		schemaXml: `<Field Type='Number' Name='RAED_MR_UCL' StaticName='RAED_MR_UCL' DisplayName='RAED_MR_UCL' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly RAEDViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"ProjectType",
	"RequirementsAnalysisEffortDensit",
	"ActualSize",
	"ActualRequirementsAnalysisEffort",
	"RAED_MR",
	"RAED_Mean",
	"RAED_LCL",
	"RAED_UCL",
	"RAED_MR_LCL",
	"RAED_MR_Mean",
	"RAED_MR_UCL"
] as const;

const definition: ListProvisionDefinition<RAEDFieldName, RAEDViewField> = {
	title: LIST_TITLE,
	description: "Requirements Analysis Effort Density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionRAED(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionRAED;
