import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
	ensureListProvision,
	FieldDefinition,
	ListProvisionDefinition,
	createLookupFieldDefinition,
	fetchListId
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.CodeReviewEffortDensity;

type CodeReviewEffortDensityFieldName =
	| "ApplicableGraphID"
	| "Goal"
	| "USL"
	| "LSL"
	| "CodeReviewEffortDensity"
	| "ActualCodeReviewEffort"
	| "ActualSize"
	| "ProjectType"
	| "CRED_LCL"
	| "CRED_Mean"
	| "CRED_UCL"
	| "CRED_MR"
	| "CRED_MR_LCL"
	| "CRED_MR_Mean"
	| "CRED_MR_UCL";

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
	},
	{
		internalName: "CRED_LCL",
		schemaXml: `<Field Type='Number' Name='CRED_LCL' StaticName='CRED_LCL' DisplayName='CRED_LCL' Decimals='2' />`
	},
	{
		internalName: "CRED_Mean",
		schemaXml: `<Field Type='Number' Name='CRED_Mean' StaticName='CRED_Mean' DisplayName='CRED_Mean' Decimals='2' />`
	},
	{
		internalName: "CRED_UCL",
		schemaXml: `<Field Type='Number' Name='CRED_UCL' StaticName='CRED_UCL' DisplayName='CRED_UCL' Decimals='2' />`
	},
	{
		internalName: "CRED_MR",
		schemaXml: `<Field Type='Number' Name='CRED_MR' StaticName='CRED_MR' DisplayName='CRED_MR' Decimals='2' />`
	},
	{
		internalName: "CRED_MR_LCL",
		schemaXml: `<Field Type='Number' Name='CRED_MR_LCL' StaticName='CRED_MR_LCL' DisplayName='CRED_MR_LCL' Decimals='2' />`
	},
	{
		internalName: "CRED_MR_Mean",
		schemaXml: `<Field Type='Number' Name='CRED_MR_Mean' StaticName='CRED_MR_Mean' DisplayName='CRED_MR_Mean' Decimals='2' />`
	},
	{
		internalName: "CRED_MR_UCL",
		schemaXml: `<Field Type='Number' Name='CRED_MR_UCL' StaticName='CRED_MR_UCL' DisplayName='CRED_MR_UCL' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly CodeReviewEffortDensityViewField[] = [
	"ApplicableGraphID",
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

export async function provisionCodeReviewEffortDensity(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
	const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
	await ensureListProvision(sp, {
		...definition,
		lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
	});
}

export default provisionCodeReviewEffortDensity;
