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

const LIST_TITLE = RequiredListsProvision.CodeReviewDefects;

type CodeReviewDefectsFieldName =
	| "CodeFileClassName"
	| "CodeFileClassSize"
	| "CodeFileAuthorDeveloperName"
	| "ReviewerName"
	| "ReviewIterationNumber"
	| "IdentifiedDate"
	| "DefectType"
	| "DefectClassification"
	| "CodeReviewChecklist"
	| "Severity"
	| "DefectStatus"
	| "ReviewResults"
	| "DefectOriginPhase"
	| "ImpactedComponents"
	| "CorrectionCorrectiveAction"
	| "PlannedClosureDate"
	| "ActualClosureDate"
	| "LocationOfDefect"
	| "Remarks";

type CodeReviewDefectsViewField = CodeReviewDefectsFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<CodeReviewDefectsFieldName>[] = [
	{
		internalName: "CodeFileClassName",
		schemaXml: `<Field Type='Text' Name='CodeFileClassName' StaticName='CodeFileClassName' DisplayName='CodeFileClassName' MaxLength='255' />`
	},
	{
		internalName: "CodeFileClassSize",
		schemaXml: `<Field Type='Text' Name='CodeFileClassSize' StaticName='CodeFileClassSize' DisplayName='CodeFileClassSize' MaxLength='255' />`
	},
	{
		internalName: "CodeFileAuthorDeveloperName",
		schemaXml: `<Field Type='User' Name='CodeFileAuthorDeveloperName' StaticName='CodeFileAuthorDeveloperName' DisplayName='CodeFileAuthorDeveloperName' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	},
	{
		internalName: "ReviewerName",
		schemaXml: `<Field Type='User' Name='ReviewerName' StaticName='ReviewerName' DisplayName='ReviewerName' UserSelectionMode='PeopleOnly' />`
	},
	{
		internalName: "ReviewIterationNumber",
		schemaXml: `<Field Type='Number' Name='ReviewIterationNumber' StaticName='ReviewIterationNumber' DisplayName='ReviewIterationNumber' />`
	},
	{
		internalName: "IdentifiedDate",
		schemaXml: `<Field Type='DateTime' Name='IdentifiedDate' StaticName='IdentifiedDate' DisplayName='IdentifiedDate' Format='DateOnly' />`
	},
	{
		internalName: "DefectType",
		schemaXml: `<Field Type='Choice' Name='DefectType' StaticName='DefectType' DisplayName='DefectType' Format='Dropdown'><CHOICES><CHOICE>Function</CHOICE><CHOICE>Interface</CHOICE><CHOICE>Checking</CHOICE><CHOICE>Assignment</CHOICE><CHOICE>Timing/ Serialization</CHOICE><CHOICE>Build/ Package</CHOICE><CHOICE>Documentation</CHOICE><CHOICE>Algorithm</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "DefectClassification",
		schemaXml: `<Field Type='Choice' Name='DefectClassification' StaticName='DefectClassification' DisplayName='DefectClassification' Format='Dropdown'><CHOICES><CHOICE>Design Issues</CHOICE><CHOICE>Data Validation</CHOICE><CHOICE>Logical</CHOICE><CHOICE>Computational</CHOICE><CHOICE>User Interface Presentation</CHOICE><CHOICE>Input/Output</CHOICE><CHOICE>Error Handling/Exception Handling</CHOICE><CHOICE>Initialization</CHOICE><CHOICE>Installation/Configuration</CHOICE><CHOICE>Performance/Load/Stress</CHOICE><CHOICE>Online help / Error messages</CHOICE><CHOICE>Third Party Software</CHOICE><CHOICE>Not Meeting Coding Standard</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "CodeReviewChecklist",
		schemaXml: `<Field Type='Text' Name='CodeReviewChecklist' StaticName='CodeReviewChecklist' DisplayName='CodeReviewChecklist' MaxLength='255' />`
	},
	{
		internalName: "Severity",
		schemaXml: `<Field Type='Choice' Name='Severity' StaticName='Severity' DisplayName='Severity' Format='Dropdown'><CHOICES><CHOICE>Critical</CHOICE><CHOICE>Moderate</CHOICE><CHOICE>Minor</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "DefectStatus",
		schemaXml: `<Field Type='Choice' Name='DefectStatus' StaticName='DefectStatus' DisplayName='DefectStatus' Format='Dropdown'><CHOICES><CHOICE>Open</CHOICE><CHOICE>Closed</CHOICE><CHOICE>Defered</CHOICE><CHOICE>In-Process</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "ReviewResults",
		schemaXml: `<Field Type='Choice' Name='ReviewResults' StaticName='ReviewResults' DisplayName='ReviewResults' Format='Dropdown'><CHOICES><CHOICE>Accepted</CHOICE><CHOICE>Rejected</CHOICE><CHOICE>Accepted with Re-Review</CHOICE><CHOICE>Rejected with Re-Review</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "DefectOriginPhase",
		schemaXml: `<Field Type='Choice' Name='DefectOriginPhase' StaticName='DefectOriginPhase' DisplayName='DefectOriginPhase' Format='Dropdown'><CHOICES><CHOICE>Requirements Specification</CHOICE><CHOICE>Designs</CHOICE><CHOICE>Code</CHOICE><CHOICE>Unit Testing</CHOICE><CHOICE>Integration Testing</CHOICE><CHOICE>System Testing</CHOICE><CHOICE>Acceptance Testing</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "ImpactedComponents",
		schemaXml: `<Field Type='Text' Name='ImpactedComponents' StaticName='ImpactedComponents' DisplayName='ImpactedComponents' MaxLength='255' />`
	},
	{
		internalName: "CorrectionCorrectiveAction",
		schemaXml: `<Field Type='Note' Name='CorrectionCorrectiveAction' StaticName='CorrectionCorrectiveAction' DisplayName='CorrectionCorrectiveAction' NumLines='6' RichText='FALSE' />`
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
		internalName: "LocationOfDefect",
		schemaXml: `<Field Type='Text' Name='LocationOfDefect' StaticName='LocationOfDefect' DisplayName='LocationOfDefect' MaxLength='255' />`
	},
	{
		internalName: "Remarks",
		schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
	}
] as const;

const defaultViewFields: readonly CodeReviewDefectsViewField[] = [
	"LinkTitle",
	"CodeFileClassName",
	"CodeFileClassSize",
	"CodeFileAuthorDeveloperName",
	"ReviewerName",
	"ReviewIterationNumber",
	"IdentifiedDate",
	"DefectType",
	"DefectClassification",
	"CodeReviewChecklist",
	"Severity",
	"DefectStatus",
	"ReviewResults",
	"DefectOriginPhase",
	"ImpactedComponents",
	"CorrectionCorrectiveAction",
	"PlannedClosureDate",
	"ActualClosureDate",
	"LocationOfDefect",
	"Remarks"
] as const;

const displayNameMappings: ReadonlyArray<{ internalName: CodeReviewDefectsFieldName | "Title"; displayName: string }> = [
	{ internalName: "Title", displayName: "Requirement ID / Ticket ID" },
	{ internalName: "CodeFileClassName", displayName: "Code File/ Class Name" },
	{ internalName: "CodeFileClassSize", displayName: "Code File/ Class Size (Lines of Code)" },
	{ internalName: "CodeFileAuthorDeveloperName", displayName: "Code File - Author/ Developer Name(s)" },
	{ internalName: "ReviewerName", displayName: "Reviewer Name" },
	{ internalName: "ReviewIterationNumber", displayName: "Review Iteration Number" },
	{ internalName: "IdentifiedDate", displayName: "Identified Date" },
	{ internalName: "DefectType", displayName: "Defect Type" },
	{ internalName: "DefectClassification", displayName: "Defect Classification" },
	{ internalName: "CodeReviewChecklist", displayName: "Code Review Checklist/ Coding Standards section ref. no." },
	{ internalName: "Severity", displayName: "Severity" },
	{ internalName: "DefectStatus", displayName: "Defect Status" },
	{ internalName: "ReviewResults", displayName: "Review Results" },
	{ internalName: "DefectOriginPhase", displayName: "Defect Origin Phase" },
	{ internalName: "ImpactedComponents", displayName: "Impacted Components" },
	{ internalName: "CorrectionCorrectiveAction", displayName: "Correction / Corrective Action" },
	{ internalName: "PlannedClosureDate", displayName: "Planned Closure Date" },
	{ internalName: "ActualClosureDate", displayName: "Actual Closure Date" },
	{ internalName: "LocationOfDefect", displayName: "Location of defect (Sub section - Line Number)" },
	{ internalName: "Remarks", displayName: "Remarks" }
];

const definition: ListProvisionDefinition<CodeReviewDefectsFieldName, CodeReviewDefectsViewField> = {
	title: LIST_TITLE,
	description: "Code review defects list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

async function ensureDefaultView(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const view = list.defaultView;
	try {
		if (typeof view.fields?.removeAll === "function") {
			await view.fields.removeAll();
		}
	} catch (error) {
		console.warn(`Failed to clear default view fields on list ${LIST_TITLE}:`, error);
	}

	for (const field of defaultViewFields) {
		try {
			await view.fields.add(field as any);
		} catch (error) {
			console.warn(`Failed to add ${field} to default view on list ${LIST_TITLE}:`, error);
		}
	}
}

async function ensureDisplayNames(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	for (const mapping of displayNameMappings) {
		try {
			await list.fields.getByInternalNameOrTitle(mapping.internalName).update({ Title: mapping.displayName });
		} catch (error) {
			console.warn(`Failed to set display name for ${mapping.internalName} on list ${LIST_TITLE}:`, error);
		}
	}
}

export async function provisionCodeReviewDefects(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureDefaultView(sp);
	await ensureDisplayNames(sp);
}

export default provisionCodeReviewDefects;