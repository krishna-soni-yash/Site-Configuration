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

const LIST_TITLE = RequiredListsProvision.CodeReviewEffortDensity;

type CodeReviewEffortDensityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "CodeReviewEffortDensity"
	| "ActualCodeReviewEffort"
	| "ActualSize"
	| "ProjectType";

type CodeReviewEffortDensityViewField = CodeReviewEffortDensityFieldName;

const fieldDefinitions: readonly FieldDefinition<CodeReviewEffortDensityFieldName>[] = [
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
		internalName: "CodeReviewEffortDensity",
		schemaXml: `<Field Type='Number' Name='CodeReviewEffortDensity' StaticName='CodeReviewEffortDensity' DisplayName='CodeReviewEffortDensity' Decimals='2' />`
	},
	{
		internalName: "ActualCodeReviewEffort",
		schemaXml: `<Field Type='Number' Name='ActualCodeReviewEffort' StaticName='ActualCodeReviewEffort' DisplayName='ActualCodeReviewEffort' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly CodeReviewEffortDensityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"CodeReviewEffortDensity",
	"ActualCodeReviewEffort",
	"ActualSize",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<CodeReviewEffortDensityFieldName, CodeReviewEffortDensityViewField> = {
	title: LIST_TITLE,
	description: "Code review effort density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionCodeReviewEffortDensity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionCodeReviewEffortDensity;
