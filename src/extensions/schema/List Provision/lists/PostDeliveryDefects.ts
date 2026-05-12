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

const LIST_TITLE = RequiredListsProvision.PostDeliveryDefects;

type PostDeliveryDefectsFieldName =
	| "Goal"
	| "USL"
	| "LSL"
	| "ProjectType"
	| "PostDeliveryDefects"
	| "ActualSize"
	| "DeliveredDefectDensity"
	| "PMDefectCount";

type PostDeliveryDefectsViewField = PostDeliveryDefectsFieldName;

const fieldDefinitions: readonly FieldDefinition<PostDeliveryDefectsFieldName>[] = [
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
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	},
	{
		internalName: "PostDeliveryDefects",
		schemaXml: `<Field Type='Number' Name='PostDeliveryDefects' StaticName='PostDeliveryDefects' DisplayName='PostDeliveryDefects' Decimals='2' />`
	},
	{
		internalName: "ActualSize",
		schemaXml: `<Field Type='Number' Name='ActualSize' StaticName='ActualSize' DisplayName='ActualSize' Decimals='2' />`
	},
	{
		internalName: "DeliveredDefectDensity",
		schemaXml: `<Field Type='Number' Name='DeliveredDefectDensity' StaticName='DeliveredDefectDensity' DisplayName='DeliveredDefectDensity' Decimals='2' />`
	},
	{
		internalName: "PMDefectCount",
		schemaXml: `<Field Type='Number' Name='PMDefectCount' StaticName='PMDefectCount' DisplayName='PMDefectCount' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly PostDeliveryDefectsViewField[] = [
	"Goal",
	"USL",
	"LSL",
	"ProjectType",
	"PostDeliveryDefects",
	"ActualSize",
	"DeliveredDefectDensity",
	"PMDefectCount"
] as const;

const definition: ListProvisionDefinition<PostDeliveryDefectsFieldName, PostDeliveryDefectsViewField> = {
	title: LIST_TITLE,
	description: "Post delivery defects",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionPostDeliveryDefects(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionPostDeliveryDefects;
