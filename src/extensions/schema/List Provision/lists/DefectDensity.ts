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

const LIST_TITLE = RequiredListsProvision.DefectDensity;

type DefectDensityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "DefectDensity"
	| "ProjectType"
	| "TestingDefects"
	| "ActualSize"
	| "PMDefectCount";

type DefectDensityViewField = DefectDensityFieldName;

const fieldDefinitions: readonly FieldDefinition<DefectDensityFieldName>[] = [
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
		internalName: "DefectDensity",
		schemaXml: `<Field Type='Number' Name='DefectDensity' StaticName='DefectDensity' DisplayName='DefectDensity' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	},
	{
		internalName: "TestingDefects",
		schemaXml: `<Field Type='Number' Name='TestingDefects' StaticName='TestingDefects' DisplayName='TestingDefects' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "PMDefectCount",
		schemaXml: `<Field Type='Number' Name='PMDefectCount' StaticName='PMDefectCount' DisplayName='PMDefectCount' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly DefectDensityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"DefectDensity",
	"ProjectType",
	"TestingDefects",
	"ActualSize",
	"PMDefectCount"
] as const;

const definition: ListProvisionDefinition<DefectDensityFieldName, DefectDensityViewField> = {
	title: LIST_TITLE,
	description: "Defect density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionDefectDensity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionDefectDensity;
