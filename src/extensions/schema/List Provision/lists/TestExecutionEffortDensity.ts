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

const LIST_TITLE = RequiredListsProvision.TestExecutionEffortDensity;

type TestExecutionEffortDensityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "ActualSize"
	| "ProjectType"
	| "TestExecutionEffortDensity"
	| "EffortSpentOnTestCasesExecution";

type TestExecutionEffortDensityViewField = TestExecutionEffortDensityFieldName;

const fieldDefinitions: readonly FieldDefinition<TestExecutionEffortDensityFieldName>[] = [
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
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	},
	{
		internalName: "TestExecutionEffortDensity",
		schemaXml: `<Field Type='Number' Name='TestExecutionEffortDensity' StaticName='TestExecutionEffortDensity' DisplayName='TestExecutionEffortDensity' Decimals='2' />`
	},
	{
		internalName: "EffortSpentOnTestCasesExecution",
		schemaXml: `<Field Type='Number' Name='EffortSpentOnTestCasesExecution' StaticName='EffortSpentOnTestCasesExecution' DisplayName='EffortSpentOnTestCasesExecution' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly TestExecutionEffortDensityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"ActualSize",
	"ProjectType",
	"TestExecutionEffortDensity",
	"EffortSpentOnTestCasesExecution"
] as const;

const definition: ListProvisionDefinition<TestExecutionEffortDensityFieldName, TestExecutionEffortDensityViewField> = {
	title: LIST_TITLE,
	description: "Test execution effort density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionTestExecutionEffortDensity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionTestExecutionEffortDensity;
