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

const LIST_TITLE = RequiredListsProvision.EffortDistribution;

type EffortDistributionFieldName =
	| "TotalEfforts"
	| "Count"
	| "ProjectType";

type EffortDistributionViewField = EffortDistributionFieldName;

const fieldDefinitions: readonly FieldDefinition<EffortDistributionFieldName>[] = [
	{
		internalName: "TotalEfforts",
		schemaXml: `<Field Type='Number' Name='TotalEfforts' StaticName='TotalEfforts' DisplayName='TotalEfforts' Decimals='2' />`
	},
	{
		internalName: "Count",
		schemaXml: `<Field Type='Number' Name='Count' StaticName='Count' DisplayName='Count' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly EffortDistributionViewField[] = [
	"TotalEfforts",
	"Count",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<EffortDistributionFieldName, EffortDistributionViewField> = {
	title: LIST_TITLE,
	description: "Effort distribution",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionEffortDistribution(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionEffortDistribution;
