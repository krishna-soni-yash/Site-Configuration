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

const LIST_TITLE = RequiredListsProvision.AMSMTTR;

type AMSMTTRFieldName =
	| "responseTimeMin"
	| "responseTimeMax"
	| "resolutionTimeMin"
	| "resolutionTimeMax";

type AMSMTTRViewField = AMSMTTRFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<AMSMTTRFieldName>[] = [
	{
		internalName: "responseTimeMin",
		schemaXml: `<Field Type='Number' Name='responseTimeMin' StaticName='responseTimeMin' DisplayName='responseTimeMin' Decimals='2' />`
	},
	{
		internalName: "responseTimeMax",
		schemaXml: `<Field Type='Number' Name='responseTimeMax' StaticName='responseTimeMax' DisplayName='responseTimeMax' Decimals='2' />`
	},
	{
		internalName: "resolutionTimeMin",
		schemaXml: `<Field Type='Number' Name='resolutionTimeMin' StaticName='resolutionTimeMin' DisplayName='resolutionTimeMin' Decimals='2' />`
	},
	{
		internalName: "resolutionTimeMax",
		schemaXml: `<Field Type='Number' Name='resolutionTimeMax' StaticName='resolutionTimeMax' DisplayName='resolutionTimeMax' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly AMSMTTRViewField[] = [
	"LinkTitle",
	"responseTimeMin",
	"responseTimeMax",
	"resolutionTimeMin",
	"resolutionTimeMax"
] as const;

const definition: ListProvisionDefinition<AMSMTTRFieldName, AMSMTTRViewField> = {
	title: LIST_TITLE,
	description: "AMSMTTR list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

async function ensureTitleRenamed(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	try {
		const field = await list.fields.getByInternalNameOrTitle("Title").select("Title")();
		if (!field || `${field.Title}` !== "priority") {
			await list.fields.getByInternalNameOrTitle("Title").update({ Title: "priority" });
		}
	} catch (error) {
		console.warn(`Failed to rename Title field on list ${LIST_TITLE}:`, error);
	}
}

export async function provisionAMSMTTR(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureTitleRenamed(sp);
}

export default provisionAMSMTTR;
