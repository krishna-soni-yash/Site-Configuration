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

const LIST_TITLE = RequiredListsProvision.FindingsSummary;

type FindingsSummaryFieldName =
	| "ApplicableGraphID"
	| "Observation"
	| "PositiveObservation"
	| "Suggestion"
	| "NC"
	| "FCFinding";

type FindingsSummaryViewField = FindingsSummaryFieldName;

const fieldDefinitions: readonly FieldDefinition<FindingsSummaryFieldName>[] = [
	{
		internalName: "Observation",
		schemaXml: `<Field Type='Number' Name='Observation' StaticName='Observation' DisplayName='Observation' Decimals='2' />`
	},
	{
		internalName: "PositiveObservation",
		schemaXml: `<Field Type='Number' Name='PositiveObservation' StaticName='PositiveObservation' DisplayName='PositiveObservation' Decimals='2' />`
	},
	{
		internalName: "Suggestion",
		schemaXml: `<Field Type='Number' Name='Suggestion' StaticName='Suggestion' DisplayName='Suggestion' Decimals='2' />`
	},
	{
		internalName: "NC",
		schemaXml: `<Field Type='Number' Name='NC' StaticName='NC' DisplayName='NC' Decimals='2' />`
	},
	{
		internalName: "FCFinding",
		schemaXml: `<Field Type='Number' Name='FCFinding' StaticName='FCFinding' DisplayName='FCFinding' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly FindingsSummaryViewField[] = [
	"ApplicableGraphID",
	"Observation",
	"PositiveObservation",
	"Suggestion",
	"NC",
	"FCFinding"
] as const;

const definition: ListProvisionDefinition<FindingsSummaryFieldName, FindingsSummaryViewField> = {
	title: LIST_TITLE,
	description: "Findings Summary",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionFindingsSummary(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
	const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
	await ensureListProvision(sp, {
		...definition,
		lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
	});
}

export default provisionFindingsSummary;
