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

const LIST_TITLE = RequiredListsProvision.AgingFindings;

type AgingFindingsFieldName =
	| "ApplicableGraphID"
	| "ObservationLe7"
	| "ObservationLe14"
	| "ObservationGt14"
	| "NCLe7"
	| "NCLe14"
	| "NCGt14"
	| "FCFindingLe7"
	| "FCFindingLe14"
	| "FCFindingGt14";

type AgingFindingsViewField = AgingFindingsFieldName;

const fieldDefinitions: readonly FieldDefinition<AgingFindingsFieldName>[] = [
	{
		internalName: "ObservationLe7",
		schemaXml: `<Field Type='Number' Name='ObservationLe7' StaticName='ObservationLe7' DisplayName='ObservationLe7' Decimals='2' />`
	},
	{
		internalName: "ObservationLe14",
		schemaXml: `<Field Type='Number' Name='ObservationLe14' StaticName='ObservationLe14' DisplayName='ObservationLe14' Decimals='2' />`
	},
	{
		internalName: "ObservationGt14",
		schemaXml: `<Field Type='Number' Name='ObservationGt14' StaticName='ObservationGt14' DisplayName='ObservationGt14' Decimals='2' />`
	},
	{
		internalName: "NCLe7",
		schemaXml: `<Field Type='Number' Name='NCLe7' StaticName='NCLe7' DisplayName='NCLe7' Decimals='2' />`
	},
	{
		internalName: "NCLe14",
		schemaXml: `<Field Type='Number' Name='NCLe14' StaticName='NCLe14' DisplayName='NCLe14' Decimals='2' />`
	},
	{
		internalName: "NCGt14",
		schemaXml: `<Field Type='Number' Name='NCGt14' StaticName='NCGt14' DisplayName='NCGt14' Decimals='2' />`
	},
	{
		internalName: "FCFindingLe7",
		schemaXml: `<Field Type='Number' Name='FCFindingLe7' StaticName='FCFindingLe7' DisplayName='FCFindingLe7' Decimals='2' />`
	},
	{
		internalName: "FCFindingLe14",
		schemaXml: `<Field Type='Number' Name='FCFindingLe14' StaticName='FCFindingLe14' DisplayName='FCFindingLe14' Decimals='2' />`
	},
	{
		internalName: "FCFindingGt14",
		schemaXml: `<Field Type='Number' Name='FCFindingGt14' StaticName='FCFindingGt14' DisplayName='FCFindingGt14' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly AgingFindingsViewField[] = [
	"ApplicableGraphID",
	"ObservationLe7",
	"ObservationLe14",
	"ObservationGt14",
	"NCLe7",
	"NCLe14",
	"NCGt14",
	"FCFindingLe7",
	"FCFindingLe14",
	"FCFindingGt14"
] as const;

const definition: ListProvisionDefinition<AgingFindingsFieldName, AgingFindingsViewField> = {
	title: LIST_TITLE,
	description: "Aging Findings",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionAgingFindings(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
	const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
	await ensureListProvision(sp, {
		...definition,
		lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
	});
}

export default provisionAgingFindings;
