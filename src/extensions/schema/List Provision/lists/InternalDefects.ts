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

const LIST_TITLE = RequiredListsProvision.InternalDefects;

type InternalDefectsFieldName =
	| "ApplicableGraphID"
	| "DefectCount"
	| "TotalEffort"
	| "ProjectType";

type InternalDefectsViewField = InternalDefectsFieldName;

const fieldDefinitions: readonly FieldDefinition<InternalDefectsFieldName>[] = [
	{
		internalName: "DefectCount",
		schemaXml: `<Field Type='Number' Name='DefectCount' StaticName='DefectCount' DisplayName='DefectCount' Decimals='2' />`
	},
	{
		internalName: "TotalEffort",
		schemaXml: `<Field Type='Number' Name='TotalEffort' StaticName='TotalEffort' DisplayName='TotalEffort' Decimals='2' />`
	},
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly InternalDefectsViewField[] = [
	"ApplicableGraphID",
	"DefectCount",
	"TotalEffort",
	"ProjectType"
] as const;

const definition: ListProvisionDefinition<InternalDefectsFieldName, InternalDefectsViewField> = {
	title: LIST_TITLE,
	description: "Internal defects",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionInternalDefects(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
	const resolvedApplicableGraphsListId = applicableGraphsListId ?? await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
	await ensureListProvision(sp, {
		...definition,
		lookupFields: [createLookupFieldDefinition("ApplicableGraphID", resolvedApplicableGraphsListId)]
	});
}

export default provisionInternalDefects;
