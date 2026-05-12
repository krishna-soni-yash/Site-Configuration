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

const LIST_TITLE = RequiredListsProvision.UnitTestingEffortDensity;

type UnitTestingEffortDensityFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "ActualSize"
	| "ProjectType"
	| "UnitTestingEffortDensity"
	| "ActualUnitTestingEffort";

type UnitTestingEffortDensityViewField = UnitTestingEffortDensityFieldName;

const fieldDefinitions: readonly FieldDefinition<UnitTestingEffortDensityFieldName>[] = [
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
		internalName: "UnitTestingEffortDensity",
		schemaXml: `<Field Type='Number' Name='UnitTestingEffortDensity' StaticName='UnitTestingEffortDensity' DisplayName='UnitTestingEffortDensity' Decimals='2' />`
	},
	{
		internalName: "ActualUnitTestingEffort",
		schemaXml: `<Field Type='Number' Name='ActualUnitTestingEffort' StaticName='ActualUnitTestingEffort' DisplayName='ActualUnitTestingEffort' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly UnitTestingEffortDensityViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"ActualSize",
	"ProjectType",
	"UnitTestingEffortDensity",
	"ActualUnitTestingEffort"
] as const;

const definition: ListProvisionDefinition<UnitTestingEffortDensityFieldName, UnitTestingEffortDensityViewField> = {
	title: LIST_TITLE,
	description: "Unit testing effort density",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionUnitTestingEffortDensity(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionUnitTestingEffortDensity;
