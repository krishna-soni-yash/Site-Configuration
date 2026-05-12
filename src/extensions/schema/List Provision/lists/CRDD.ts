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

const LIST_TITLE = RequiredListsProvision.CRDD;

type CRDDFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "ProjectType"
	| "CodeReviewDefectDensity"
	| "ActualCodeReviewDefects"
	| "ActualSize"
	| "CRDD_MR"
	| "CRDD_LCL"
	| "CRDD_Mean"
	| "CRDD_UCL"
	| "CRDD_MR_LCL"
	| "CRDD_MR_Mean"
	| "CRDD_MR_UCL";

type CRDDViewField = CRDDFieldName;

const fieldDefinitions: readonly FieldDefinition<CRDDFieldName>[] = [
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
		internalName: "CodeReviewDefectDensity",
		schemaXml: `<Field Type='Number' Name='CodeReviewDefectDensity' StaticName='CodeReviewDefectDensity' DisplayName='CodeReviewDefectDensity' Decimals='2' />`
	},
	{
		internalName: "ActualCodeReviewDefects",
		schemaXml: `<Field Type='Number' Name='ActualCodeReviewDefects' StaticName='ActualCodeReviewDefects' DisplayName='ActualCodeReviewDefects' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "CRDD_MR",
		schemaXml: `<Field Type='Number' Name='CRDD_MR' StaticName='CRDD_MR' DisplayName='CRDD_MR' Decimals='2' />`
	},
	{
		internalName: "CRDD_LCL",
		schemaXml: `<Field Type='Number' Name='CRDD_LCL' StaticName='CRDD_LCL' DisplayName='CRDD_LCL' Decimals='2' />`
	},
	{
		internalName: "CRDD_Mean",
		schemaXml: `<Field Type='Number' Name='CRDD_Mean' StaticName='CRDD_Mean' DisplayName='CRDD_Mean' Decimals='2' />`
	},
	{
		internalName: "CRDD_UCL",
		schemaXml: `<Field Type='Number' Name='CRDD_UCL' StaticName='CRDD_UCL' DisplayName='CRDD_UCL' Decimals='2' />`
	},
	{
		internalName: "CRDD_MR_LCL",
		schemaXml: `<Field Type='Number' Name='CRDD_MR_LCL' StaticName='CRDD_MR_LCL' DisplayName='CRDD_MR_LCL' Decimals='2' />`
	},
	{
		internalName: "CRDD_MR_Mean",
		schemaXml: `<Field Type='Number' Name='CRDD_MR_Mean' StaticName='CRDD_MR_Mean' DisplayName='CRDD_MR_Mean' Decimals='2' />`
	},
	{
		internalName: "CRDD_MR_UCL",
		schemaXml: `<Field Type='Number' Name='CRDD_MR_UCL' StaticName='CRDD_MR_UCL' DisplayName='CRDD_MR_UCL' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly CRDDViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"ProjectType",
	"CodeReviewDefectDensity",
	"ActualCodeReviewDefects",
	"ActualSize",
	"CRDD_MR",
	"CRDD_LCL",
	"CRDD_Mean",
	"CRDD_UCL",
	"CRDD_MR_LCL",
	"CRDD_MR_Mean",
	"CRDD_MR_UCL"
] as const;

const definition: ListProvisionDefinition<CRDDFieldName, CRDDViewField> = {
	title: LIST_TITLE,
	description: "Code review defect density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionCRDD(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionCRDD;
