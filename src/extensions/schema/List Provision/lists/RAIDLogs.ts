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

const LIST_TITLE = RequiredListsProvision.RAIDLogs;

type RAIDLogsFieldName =
	| "SelectType"
	| "RAIDId"
	| "TypeOfAction"
	| "ByWhom"
	| "Responsibility"
	| "Remarks"
	| "IdentificationDate"
	| "RiskDescription"
	| "Impact"
	| "RiskPriority"
	| "ActionPlan"
	| "RiskStatus"
	| "RiskExposure";

type RAIDLogsViewField = RAIDLogsFieldName | "ID" | "Modified" | "Editor";

const selectTypeChoices = ["Issue", "Assumption", "Dependency", "Risk"] as const;
const typeOfActionChoices = ["Mitigation", "Contingency", "Leverage"] as const;
const riskPriorityChoices = ["High", "Medium", "Low"] as const;
const riskStatusChoices = ["Closed", "In Progress", "Monitoring"] as const;

function buildChoiceFieldSchema(name: string, displayName: string, choices: readonly string[]): string {
	const choicesXml = choices.map((choice) => `<CHOICE>${choice}</CHOICE>`).join("");
	return `<Field Type='Choice' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='Dropdown'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

const fieldDefinitions: readonly FieldDefinition<RAIDLogsFieldName>[] = [
	{
		internalName: "SelectType",
		schemaXml: buildChoiceFieldSchema("SelectType", "SelectType", selectTypeChoices)
	},
	{
		internalName: "RAIDId",
		schemaXml: `<Field Type='Text' Name='RAIDId' StaticName='RAIDId' DisplayName='RAIDId' MaxLength='255' />`
	},
	{
		internalName: "TypeOfAction",
		schemaXml: buildChoiceFieldSchema("TypeOfAction", "TypeOfAction", typeOfActionChoices)
	},
	{
		internalName: "ByWhom",
		schemaXml: `<Field Type='User' Name='ByWhom' StaticName='ByWhom' DisplayName='ByWhom' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	},
	{
		internalName: "Responsibility",
		schemaXml: `<Field Type='User' Name='Responsibility' StaticName='Responsibility' DisplayName='Responsibility' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	},
	{
		internalName: "Remarks",
		schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "IdentificationDate",
		schemaXml: `<Field Type='DateTime' Name='IdentificationDate' StaticName='IdentificationDate' DisplayName='IdentificationDate' Format='DateOnly' />`
	},
	{
		internalName: "RiskDescription",
		schemaXml: `<Field Type='Note' Name='RiskDescription' StaticName='RiskDescription' DisplayName='RiskDescription' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "Impact",
		schemaXml: `<Field Type='Text' Name='Impact' StaticName='Impact' DisplayName='Impact' MaxLength='255' />`
	},
	{
		internalName: "RiskPriority",
		schemaXml: buildChoiceFieldSchema("RiskPriority", "RiskPriority", riskPriorityChoices)
	},
	{
		internalName: "ActionPlan",
		schemaXml: `<Field Type='Note' Name='ActionPlan' StaticName='ActionPlan' DisplayName='ActionPlan' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "RiskStatus",
		schemaXml: buildChoiceFieldSchema("RiskStatus", "RiskStatus", riskStatusChoices)
	},
	{
		internalName: "RiskExposure",
		schemaXml: `<Field Type='Number' Name='RiskExposure' StaticName='RiskExposure' DisplayName='RiskExposure' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly RAIDLogsViewField[] = [
	"ID",
	"RAIDId",
	"SelectType",
	"TypeOfAction",
	"ByWhom",
	"Responsibility",
	"Modified",
	"Editor",
	"Remarks",
	"IdentificationDate",
	"RiskDescription",
	"Impact",
	"RiskPriority",
	"ActionPlan",
	"RiskStatus",
	"RiskExposure"
] as const;

const definition: ListProvisionDefinition<RAIDLogsFieldName, RAIDLogsViewField> = {
	title: LIST_TITLE,
	description: "RAID logs list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionRAIDLogs(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionRAIDLogs;
