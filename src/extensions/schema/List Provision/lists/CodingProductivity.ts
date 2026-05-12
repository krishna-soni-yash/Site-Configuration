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

const LIST_TITLE = RequiredListsProvision.CodingProductivity;

type CodingProductivityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "CodingProductivity"
	| "CP_MR"
	| "CodingEffort"
	| "ActualSize"
	| "CP_LCL"
	| "CP_UCL"
	| "CP_Mean"
	| "CP_MR_LCL"
	| "CP_MR_UCL"
	| "CP_MR_Mean"
	| "ProjectType";

type CodingProductivityViewField = CodingProductivityFieldName;

const fieldDefinitions: readonly FieldDefinition<CodingProductivityFieldName>[] = [
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
		internalName: "CodingProductivity",
		schemaXml: `<Field Type='Number' Name='CodingProductivity' StaticName='CodingProductivity' DisplayName='CodingProductivity' Decimals='2' />`
	},
	{
		internalName: "CP_MR",
		schemaXml: `<Field Type='Number' Name='CP_MR' StaticName='CP_MR' DisplayName='CP_MR' Decimals='2' />`
	},
	{
		internalName: "CodingEffort",
		schemaXml: `<Field Type='Number' Name='CodingEffort' StaticName='CodingEffort' DisplayName='CodingEffort' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "CP_LCL",
		schemaXml: `<Field Type='Number' Name='CP_LCL' StaticName='CP_LCL' DisplayName='CP_LCL' Decimals='2' />`
	},
	{
		internalName: "CP_UCL",
		schemaXml: `<Field Type='Number' Name='CP_UCL' StaticName='CP_UCL' DisplayName='CP_UCL' Decimals='2' />`
	},
	{
		internalName: "CP_Mean",
		schemaXml: `<Field Type='Number' Name='CP_Mean' StaticName='CP_Mean' DisplayName='CP_Mean' Decimals='2' />`
	},
	{
		internalName: "CP_MR_LCL",
		schemaXml: `<Field Type='Number' Name='CP_MR_LCL' StaticName='CP_MR_LCL' DisplayName='CP_MR_LCL' Decimals='2' />`
	},
	{
		internalName: "CP_MR_UCL",
		schemaXml: `<Field Type='Number' Name='CP_MR_UCL' StaticName='CP_MR_UCL' DisplayName='CP_MR_UCL' Decimals='2' />`
	},
	{
		internalName: "CP_MR_Mean",
		schemaXml: `<Field Type='Number' Name='CP_MR_Mean' StaticName='CP_MR_Mean' DisplayName='CP_MR_Mean' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly CodingProductivityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"CodingProductivity",
	"CP_MR",
	"CodingEffort",
	"ActualSize",
	"CP_LCL",
	"CP_UCL",
	"CP_Mean",
	"CP_MR_LCL",
	"CP_MR_UCL",
	"CP_MR_Mean",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<CodingProductivityFieldName, CodingProductivityViewField> = {
	title: LIST_TITLE,
	description: "Coding productivity",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionCodingProductivity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionCodingProductivity;
