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

const LIST_TITLE = RequiredListsProvision.ComplexityWeightage;

type ComplexityWeightageFieldName = "Simple" | "Medium" | "Complex" | "VeryComplex";

type ComplexityWeightageViewField = ComplexityWeightageFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<ComplexityWeightageFieldName>[] = [
	{
		internalName: "Simple",
		schemaXml: `<Field Type='Text' Name='Simple' StaticName='Simple' DisplayName='Simple' MaxLength='255' />`
	},
	{
		internalName: "Medium",
		schemaXml: `<Field Type='Text' Name='Medium' StaticName='Medium' DisplayName='Medium' MaxLength='255' />`
	},
	{
		internalName: "Complex",
		schemaXml: `<Field Type='Text' Name='Complex' StaticName='Complex' DisplayName='Complex' MaxLength='255' />`
	},
	{
		internalName: "VeryComplex",
		schemaXml: `<Field Type='Text' Name='VeryComplex' StaticName='VeryComplex' DisplayName='VeryComplex' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly ComplexityWeightageViewField[] = [
	"LinkTitle",
	"Simple",
	"Medium",
	"Complex",
	"VeryComplex"
] as const;

const definition: ListProvisionDefinition<ComplexityWeightageFieldName, ComplexityWeightageViewField> = {
	title: LIST_TITLE,
	description: "Complexity weightage list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

async function ensureTitleRenamed(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	try {
		const field = await list.fields.getByInternalNameOrTitle("Title").select("Title")();
		if (!field || `${field.Title}` !== "ProjectType") {
			await list.fields.getByInternalNameOrTitle("Title").update({ Title: "ProjectType" });
		}
	} catch (error) {
		console.warn(`Failed to rename Title field on list ${LIST_TITLE}:`, error);
	}
}

export async function provisionComplexityWeightage(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureTitleRenamed(sp);
}

export default provisionComplexityWeightage;
