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

const LIST_TITLE = RequiredListsProvision.AdjustmentFactorValue;

type AdjustmentFactorValueFieldName =
	| "ProjectType"
	| "VeryHigh"
	| "High"
	| "Medium"
	| "Low"
	| "VeryLow"
	| "None"
	| "NoReuse";

type AdjustmentFactorValueViewField = AdjustmentFactorValueFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<AdjustmentFactorValueFieldName>[] = [
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	},
	{
		internalName: "VeryHigh",
		schemaXml: `<Field Type='Text' Name='VeryHigh' StaticName='VeryHigh' DisplayName='VeryHigh' MaxLength='255' />`
	},
	{
		internalName: "High",
		schemaXml: `<Field Type='Text' Name='High' StaticName='High' DisplayName='High' MaxLength='255' />`
	},
	{
		internalName: "Medium",
		schemaXml: `<Field Type='Text' Name='Medium' StaticName='Medium' DisplayName='Medium' MaxLength='255' />`
	},
	{
		internalName: "Low",
		schemaXml: `<Field Type='Text' Name='Low' StaticName='Low' DisplayName='Low' MaxLength='255' />`
	},
	{
		internalName: "VeryLow",
		schemaXml: `<Field Type='Text' Name='VeryLow' StaticName='VeryLow' DisplayName='VeryLow' MaxLength='255' />`
	},
	{
		internalName: "None",
		schemaXml: `<Field Type='Text' Name='None' StaticName='None' DisplayName='None' MaxLength='255' />`
	},
	{
		internalName: "NoReuse",
		schemaXml: `<Field Type='Text' Name='NoReuse' StaticName='NoReuse' DisplayName='NoReuse' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly AdjustmentFactorValueViewField[] = [
	"LinkTitle",
	"ProjectType",
	"VeryHigh",
	"High",
	"Medium",
	"Low",
	"VeryLow",
	"None",
	"NoReuse"
] as const;

const definition: ListProvisionDefinition<AdjustmentFactorValueFieldName, AdjustmentFactorValueViewField> = {
	title: LIST_TITLE,
	description: "Adjustment factor value list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

async function ensureTitleRenamed(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	try {
		const field = await list.fields.getByInternalNameOrTitle("Title").select("Title")();
		if (!field || `${field.Title}` !== "FactorName") {
			await list.fields.getByInternalNameOrTitle("Title").update({ Title: "FactorName" });
		}
	} catch (error) {
		console.warn(`Failed to rename Title field on list ${LIST_TITLE}:`, error);
	}
}

export async function provisionAdjustmentFactorValue(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureTitleRenamed(sp);
}

export default provisionAdjustmentFactorValue;
