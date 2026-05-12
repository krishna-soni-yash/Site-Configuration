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

const LIST_TITLE = RequiredListsProvision.EffortVariation;

type EffortVariationFieldName =
	| "PlannedEffort"
	| "ActualEffort"
	| "Goal"
	| "USL"
	| "LSL"
	| "EffortVariation"
	| "Mean"
	| "Variance"
	| "ProjectType";

type EffortVariationViewField = EffortVariationFieldName;

const fieldDefinitions: readonly FieldDefinition<EffortVariationFieldName>[] = [
	{
		internalName: "PlannedEffort",
		schemaXml: `<Field Type='Number' Name='PlannedEffort' StaticName='PlannedEffort' DisplayName='PlannedEffort' Decimals='2' />`
	},
	{
		internalName: "ActualEffort",
		schemaXml: `<Field Type='Number' Name='ActualEffort' StaticName='ActualEffort' DisplayName='ActualEffort' Decimals='2' />`
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
		internalName: "EffortVariation",
		schemaXml: `<Field Type='Number' Name='EffortVariation' StaticName='EffortVariation' DisplayName='EffortVariation' Decimals='2' />`
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

const defaultViewFields: readonly EffortVariationViewField[] = [
	"PlannedEffort",
	"ActualEffort",
	"Goal",
	"USL",
	"LSL",
	"EffortVariation",
	"Mean",
	"Variance",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<EffortVariationFieldName, EffortVariationViewField> = {
	title: LIST_TITLE,
	description: "Effort variation",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionEffortVariation(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionEffortVariation;
