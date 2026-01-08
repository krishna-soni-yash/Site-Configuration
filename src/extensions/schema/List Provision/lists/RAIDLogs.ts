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
	| "RiskExposure"
	| "IssueDetails"
	| "IDADate"
	| "ImplementationActions"
	| "PlannedClosureDate"
	| "ActualClosureDate"
	| "Applicability"
	| "AssociatedGoal"
	| "RiskSource"
	| "RiskCategory"
	| "PotentialCost"
	| "PotentialBenefit"
	| "TargetDate"
	| "ActualDate"
	| "Effectiveness"
	| "ImpactValue"
	| "ProbabilityValue"
	| "OpportunityValue";

type RAIDLogsViewField = RAIDLogsFieldName | "ID" | "Modified" | "Editor";

const selectTypeChoices = ["Issue", "Assumption", "Dependency", "Risk"] as const;
const typeOfActionChoices = ["Mitigation", "Contingency", "Leverage"] as const;
const riskPriorityChoices = ["High", "Medium", "Low"] as const;
const riskStatusChoices = ["Closed", "In Progress", "Monitoring"] as const;
const applicabilityChoices = ["Yes", "No"] as const;
const associatedGoalChoices = ["BG01", "BG02"] as const;
const riskSourceChoices = ["Internal", "External"] as const;
const riskCategoryChoices = ["Resource", "Customer", "Information Security", "Technology", "Business", "Process", "Others"] as const;

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
	},
	{
		internalName: "IssueDetails",
		schemaXml: `<Field Type='Note' Name='IssueDetails' StaticName='IssueDetails' DisplayName='IssueDetails' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "IDADate",
		schemaXml: `<Field Type='DateTime' Name='IDADate' StaticName='IDADate' DisplayName='IDADate' Format='DateOnly' />`
	},
	{
		internalName: "ImplementationActions",
		schemaXml: `<Field Type='Note' Name='ImplementationActions' StaticName='ImplementationActions' DisplayName='ImplementationActions' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "PlannedClosureDate",
		schemaXml: `<Field Type='DateTime' Name='PlannedClosureDate' StaticName='PlannedClosureDate' DisplayName='PlannedClosureDate' Format='DateOnly' />`
	},
	{
		internalName: "ActualClosureDate",
		schemaXml: `<Field Type='DateTime' Name='ActualClosureDate' StaticName='ActualClosureDate' DisplayName='ActualClosureDate' Format='DateOnly' />`
	},
	{
		internalName: "Applicability",
		schemaXml: buildChoiceFieldSchema("Applicability", "Applicability", applicabilityChoices)
	},
	{
		internalName: "AssociatedGoal",
		schemaXml: buildChoiceFieldSchema("AssociatedGoal", "AssociatedGoal", associatedGoalChoices)
	},
	{
		internalName: "RiskSource",
		schemaXml: buildChoiceFieldSchema("RiskSource", "RiskSource", riskSourceChoices)
	},
	{
		internalName: "RiskCategory",
		schemaXml: buildChoiceFieldSchema("RiskCategory", "RiskCategory", riskCategoryChoices)
	},
	{
		internalName: "PotentialCost",
		schemaXml: `<Field Type='Number' Name='PotentialCost' StaticName='PotentialCost' DisplayName='PotentialCost' Decimals='2' />`
	},
	{
		internalName: "PotentialBenefit",
		schemaXml: `<Field Type='Number' Name='PotentialBenefit' StaticName='PotentialBenefit' DisplayName='PotentialBenefit' Decimals='2' />`
	},
	{
		internalName: "TargetDate",
		schemaXml: `<Field Type='DateTime' Name='TargetDate' StaticName='TargetDate' DisplayName='TargetDate' Format='DateOnly' />`
	},
	{
		internalName: "ActualDate",
		schemaXml: `<Field Type='DateTime' Name='ActualDate' StaticName='ActualDate' DisplayName='ActualDate' Format='DateOnly' />`
	},
	{
		internalName: "Effectiveness",
		schemaXml: `<Field Type='Note' Name='Effectiveness' StaticName='Effectiveness' DisplayName='Effectiveness' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "ImpactValue",
		schemaXml: `<Field Type='Number' Name='ImpactValue' StaticName='ImpactValue' DisplayName='ImpactValue' Decimals='0' />`
	},
	{
		internalName: "ProbabilityValue",
		schemaXml: `<Field Type='Number' Name='ProbabilityValue' StaticName='ProbabilityValue' DisplayName='ProbabilityValue' Decimals='0' />`
	},
	{
		internalName: "OpportunityValue",
		schemaXml: `<Field Type='Number' Name='OpportunityValue' StaticName='OpportunityValue' DisplayName='OpportunityValue' Decimals='2' />`
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
	"RiskExposure",
	"IssueDetails",
	"IDADate",
	"ImplementationActions",
	"PlannedClosureDate",
	"ActualClosureDate",
	"Applicability",
	"AssociatedGoal",
	"RiskSource",
	"RiskCategory",
	"PotentialCost",
	"PotentialBenefit",
	"TargetDate",
	"ActualDate",
	"Effectiveness",
	"ImpactValue",
	"ProbabilityValue",
	"OpportunityValue"
] as const;

const fieldDisplayNameUpdates: Partial<Record<RAIDLogsFieldName, string>> = {
	IssueDetails: "Issue/Assumption/Dependencies Details",
	IDADate: "Date",
	ImplementationActions: "Implementation Actions",
	PlannedClosureDate: "Planned Closure Date",
	ActualClosureDate: "Actual Closure Date",
	AssociatedGoal: "Associated Goal",
	RiskSource: "Source",
	RiskCategory: "Category",
	PotentialCost: "Potential Cost",
	PotentialBenefit: "Potential Benefit",
	TargetDate: "Target Date",
	ActualDate: "Actual Date",
	ImpactValue: "Impact Value (1 to 10)",
	ProbabilityValue: "Probability Value(1 to 10)",
	OpportunityValue: "Opportunity Value"
};

async function applyFieldDisplayNames(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	for (const [internalName, displayName] of Object.entries(fieldDisplayNameUpdates)) {
		try {
			await list.fields.getByInternalNameOrTitle(internalName).update({ Title: displayName });
		} catch (error) {
			console.warn(`Failed to update display name for ${internalName} on ${LIST_TITLE}:`, error);
		}
	}
}

const definition: ListProvisionDefinition<RAIDLogsFieldName, RAIDLogsViewField> = {
	title: LIST_TITLE,
	description: "RAID logs list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionRAIDLogs(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await applyFieldDisplayNames(sp);
}

export default provisionRAIDLogs;
