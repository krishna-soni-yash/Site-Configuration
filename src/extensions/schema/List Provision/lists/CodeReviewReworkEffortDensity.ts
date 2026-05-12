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

const LIST_TITLE = RequiredListsProvision.CodeReviewReworkEffortDensity;

type CodeReviewReworkEffortDensityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "CodeReviewReworkEffortDensity"
	| "ActualSize"
	| "CodeReviewReworkEffort"
	| "ProjectType";

type CodeReviewReworkEffortDensityViewField = CodeReviewReworkEffortDensityFieldName;

const fieldDefinitions: readonly FieldDefinition<CodeReviewReworkEffortDensityFieldName>[] = [
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
		internalName: "CodeReviewReworkEffortDensity",
		schemaXml: `<Field Type='Number' Name='CodeReviewReworkEffortDensity' StaticName='CodeReviewReworkEffortDensity' DisplayName='CodeReviewReworkEffortDensity' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "CodeReviewReworkEffort",
		schemaXml: `<Field Type='Number' Name='CodeReviewReworkEffort' StaticName='CodeReviewReworkEffort' DisplayName='CodeReviewReworkEffort' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly CodeReviewReworkEffortDensityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"CodeReviewReworkEffortDensity",
	"ActualSize",
	"CodeReviewReworkEffort",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<CodeReviewReworkEffortDensityFieldName, CodeReviewReworkEffortDensityViewField> = {
	title: LIST_TITLE,
	description: "Code review rework effort density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionCodeReviewReworkEffortDensity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionCodeReviewReworkEffortDensity;
