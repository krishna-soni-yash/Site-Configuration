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

const LIST_TITLE = RequiredListsProvision.LlBpRc;

type LlBpRcFieldName =
	| "DataType"
	| "LlProblemFacedLearning"
	| "LlCategory"
	| "LlSolution"
	| "LlRemarks"
	| "BpBestPracticesDescription"
	| "BpCategory"
	| "BpRemarks"
	| "RcComponentName"
	| "RcLocation"
	| "RcPurposeMainFunctionality"
	| "RcRemarks"
	| "Attachments";

type LlBpRcViewField = LlBpRcFieldName;

const fieldDefinitions: readonly FieldDefinition<LlBpRcFieldName>[] = [
	{
		internalName: "DataType",
		schemaXml: `<Field Type='Text' Name='DataType' StaticName='DataType' DisplayName='DataType' MaxLength='255' />`
	},
	{
		internalName: "LlProblemFacedLearning",
		schemaXml: `<Field Type='Note' Name='LlProblemFacedLearning' StaticName='LlProblemFacedLearning' DisplayName='LlProblemFacedLearning' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "LlCategory",
		schemaXml: `<Field Type='Note' Name='LlCategory' StaticName='LlCategory' DisplayName='LlCategory' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "LlSolution",
		schemaXml: `<Field Type='Note' Name='LlSolution' StaticName='LlSolution' DisplayName='LlSolution' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "LlRemarks",
		schemaXml: `<Field Type='Note' Name='LlRemarks' StaticName='LlRemarks' DisplayName='LlRemarks' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "BpBestPracticesDescription",
		schemaXml: `<Field Type='Note' Name='BpBestPracticesDescription' StaticName='BpBestPracticesDescription' DisplayName='BpBestPracticesDescription' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "BpCategory",
		schemaXml: `<Field Type='Note' Name='BpCategory' StaticName='BpCategory' DisplayName='BpCategory' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "BpRemarks",
		schemaXml: `<Field Type='Note' Name='BpRemarks' StaticName='BpRemarks' DisplayName='BpRemarks' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "RcComponentName",
		schemaXml: `<Field Type='Note' Name='RcComponentName' StaticName='RcComponentName' DisplayName='RcComponentName' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "RcLocation",
		schemaXml: `<Field Type='Text' Name='RcLocation' StaticName='RcLocation' DisplayName='RcLocation' MaxLength='255' />`
	},
	{
		internalName: "RcPurposeMainFunctionality",
		schemaXml: `<Field Type='Note' Name='RcPurposeMainFunctionality' StaticName='RcPurposeMainFunctionality' DisplayName='RcPurposeMainFunctionality' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "RcRemarks",
		schemaXml: `<Field Type='Note' Name='RcRemarks' StaticName='RcRemarks' DisplayName='RcRemarks' NumLines='6' RichText='FALSE' />`
	}
] as const;

const lessonsLearntViewFields: readonly LlBpRcViewField[] = [
	"DataType",
	"LlProblemFacedLearning",
	"LlCategory",
	"LlSolution",
	"LlRemarks",
	"Attachments"
] as const;

const bestPracticesViewFields: readonly LlBpRcViewField[] = [
	"DataType",
	"BpBestPracticesDescription",
	"BpCategory",
	"BpRemarks",
	"Attachments"
] as const;

const reusableComponentsViewFields: readonly LlBpRcViewField[] = [
	"DataType",
	"RcComponentName",
	"RcLocation",
	"RcPurposeMainFunctionality",
	"RcRemarks",
	"Attachments"
] as const;

const viewDefinitions: ReadonlyArray<{
	title: string;
	fields: readonly LlBpRcViewField[];
	makeDefault?: boolean;
	includeLinkTitle?: boolean;
	query?: string;
}> = [
	{
		title: "Lessons Learnt",
		fields: lessonsLearntViewFields,
		makeDefault: true,
		includeLinkTitle: false,
		query: `<Where><Eq><FieldRef Name='DataType' /><Value Type='Text'>LessonsLearnt</Value></Eq></Where>`
	},
	{
		title: "Best Practices",
		fields: bestPracticesViewFields,
		includeLinkTitle: false,
		query: `<Where><Eq><FieldRef Name='DataType' /><Value Type='Text'>BestPractices</Value></Eq></Where>`
	},
	{
		title: "Reusable Components",
		fields: reusableComponentsViewFields,
		includeLinkTitle: false,
		query: `<Where><Eq><FieldRef Name='DataType' /><Value Type='Text'>ReusableComponents</Value></Eq></Where>`
	}
];

const definition: ListProvisionDefinition<LlBpRcFieldName, LlBpRcViewField> = {
	title: LIST_TITLE,
	description: "Lessons learnt, best practices, and reusable components",
	templateId: 100,
	fields: fieldDefinitions,
	indexedFields: [],
	defaultViewFields: lessonsLearntViewFields,
	views: viewDefinitions
};

export async function provisionLlBpRc(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionLlBpRc;
