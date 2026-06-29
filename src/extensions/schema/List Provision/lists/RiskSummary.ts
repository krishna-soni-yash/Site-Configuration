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

const LIST_TITLE = RequiredListsProvision.RiskSummary;

type RiskSummaryFieldName =
	| "REGreaterThanEqToEighty"
	| "REGreaterThanEqToSixty"
	| "REGreaterThanEqToZero";

type RiskSummaryViewField = RiskSummaryFieldName;

const fieldDefinitions: readonly FieldDefinition<RiskSummaryFieldName>[] = [
	{
		internalName: "REGreaterThanEqToEighty",
		schemaXml: `<Field Type='Number' Name='REGreaterThanEqToEighty' StaticName='REGreaterThanEqToEighty' DisplayName='REGreaterThanEqToEighty' Decimals='2' />`
	},
	{
		internalName: "REGreaterThanEqToSixty",
		schemaXml: `<Field Type='Number' Name='REGreaterThanEqToSixty' StaticName='REGreaterThanEqToSixty' DisplayName='REGreaterThanEqToSixty' Decimals='2' />`
	},
	{
		internalName: "REGreaterThanEqToZero",
		schemaXml: `<Field Type='Number' Name='REGreaterThanEqToZero' StaticName='REGreaterThanEqToZero' DisplayName='REGreaterThanEqToZero' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly RiskSummaryViewField[] = [
	"REGreaterThanEqToEighty",
	"REGreaterThanEqToSixty",
	"REGreaterThanEqToZero"
] as const;

const definition: ListProvisionDefinition<RiskSummaryFieldName, RiskSummaryViewField> = {
	title: LIST_TITLE,
	description: "Risk Summary",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionRiskSummary(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionRiskSummary;
