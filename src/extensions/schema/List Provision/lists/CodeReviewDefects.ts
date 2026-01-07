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
		schemaXml: `<Field Type='Text' Name='CodeFileClassName' StaticName='CodeFileClassName' DisplayName='Code File/ Class Name' MaxLength='255' />`
	},
	{
		internalName: "CodeFileClassSize",
		schemaXml: `<Field Type='Text' Name='CodeFileClassSize' StaticName='CodeFileClassSize' DisplayName='Code File/ Class Size (Lines of Code)' MaxLength='255' />`
	},
	{
		internalName: "CodeFileAuthorDeveloperName",
		schemaXml: `<Field Type='User' Name='CodeFileAuthorDeveloperName' StaticName='CodeFileAuthorDeveloperName' DisplayName='Code File - Author/ Developer Name(s)' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	},
	{
		internalName: "ReviewerName",
		schemaXml: `<Field Type='User' Name='ReviewerName' StaticName='ReviewerName' DisplayName='Reviewer Name' UserSelectionMode='PeopleOnly' />`
	},
	{
		internalName: "ReviewIterationNumber",
		schemaXml: `<Field Type='Number' Name='ReviewIterationNumber' StaticName='ReviewIterationNumber' DisplayName='Review Iteration Number' />`
	},
	{
		internalName: "IdentifiedDate",
		schemaXml: `<Field Type='DateTime' Name='IdentifiedDate' StaticName='IdentifiedDate' DisplayName='Identified Date' Format='DateOnly' />`
	},
	{
		internalName: "DefectType",
		schemaXml: `<Field Type='Choice' Name='DefectType' StaticName='DefectType' DisplayName='Defect Type' Format='Dropdown'><CHOICES><CHOICE>Function</CHOICE><CHOICE>Interface</CHOICE><CHOICE>Checking</CHOICE><CHOICE>Assignment</CHOICE><CHOICE>Timing/ Serialization</CHOICE><CHOICE>Build/ Package</CHOICE><CHOICE>Documentation</CHOICE><CHOICE>Algorithm</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "DefectClassification",
		schemaXml: `<Field Type='Choice' Name='DefectClassification' StaticName='DefectClassification' DisplayName='Defect Classification' Format='Dropdown'><CHOICES><CHOICE>Design Issues</CHOICE><CHOICE>Data Validation</CHOICE><CHOICE>Logical</CHOICE><CHOICE>Computational</CHOICE><CHOICE>User Interface Presentation</CHOICE><CHOICE>Input/Output</CHOICE><CHOICE>Error Handling/Exception Handling</CHOICE><CHOICE>Initialization</CHOICE><CHOICE>Installation/Configuration</CHOICE><CHOICE>Performance/Load/Stress</CHOICE><CHOICE>Online help / Error messages</CHOICE><CHOICE>Third Party Software</CHOICE><CHOICE>Not Meeting Coding Standard</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "CodeReviewChecklist",
		schemaXml: `<Field Type='Text' Name='CodeReviewChecklist' StaticName='CodeReviewChecklist' DisplayName='Code Review Checklist/ Coding Standards section ref. no.' MaxLength='255' />`
	},
	{
		internalName: "Severity",
		schemaXml: `<Field Type='Choice' Name='Severity' StaticName='Severity' DisplayName='Severity' Format='Dropdown'><CHOICES><CHOICE>Critical</CHOICE><CHOICE>Moderate</CHOICE><CHOICE>Minor</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "DefectStatus",
		schemaXml: `<Field Type='Choice' Name='DefectStatus' StaticName='DefectStatus' DisplayName='Defect Status' Format='Dropdown'><CHOICES><CHOICE>Open</CHOICE><CHOICE>Closed</CHOICE><CHOICE>Defered</CHOICE><CHOICE>In-Process</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "ReviewResults",
		schemaXml: `<Field Type='Choice' Name='ReviewResults' StaticName='ReviewResults' DisplayName='Review Results' Format='Dropdown'><CHOICES><CHOICE>Accepted</CHOICE><CHOICE>Rejected</CHOICE><CHOICE>Accepted with Re-Review</CHOICE><CHOICE>Rejected with Re-Review</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "DefectOriginPhase",
		schemaXml: `<Field Type='Choice' Name='DefectOriginPhase' StaticName='DefectOriginPhase' DisplayName='Defect Origin Phase' Format='Dropdown'><CHOICES><CHOICE>Requirements Specification</CHOICE><CHOICE>Designs</CHOICE><CHOICE>Code</CHOICE><CHOICE>Unit Testing</CHOICE><CHOICE>Integration Testing</CHOICE><CHOICE>System Testing</CHOICE><CHOICE>Acceptance Testing</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "ImpactedComponents",
		schemaXml: `<Field Type='Text' Name='ImpactedComponents' StaticName='ImpactedComponents' DisplayName='Impacted Components' MaxLength='255' />`
	},
	{
		internalName: "CorrectionCorrectiveAction",
		schemaXml: `<Field Type='Note' Name='CorrectionCorrectiveAction' StaticName='CorrectionCorrectiveAction' DisplayName='Correction / Corrective Action' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "PlannedClosureDate",
		schemaXml: `<Field Type='DateTime' Name='PlannedClosureDate' StaticName='PlannedClosureDate' DisplayName='Planned Closure Date' Format='DateOnly' />`
	},
	{
		internalName: "ActualClosureDate",
		schemaXml: `<Field Type='DateTime' Name='ActualClosureDate' StaticName='ActualClosureDate' DisplayName='Actual Closure Date' Format='DateOnly' />`
	},
	{
		internalName: "LocationOfDefect",
		schemaXml: `<Field Type='Text' Name='LocationOfDefect' StaticName='LocationOfDefect' DisplayName='Location of defect (Sub section - Line Number)' MaxLength='255' />`
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

const definition: ListProvisionDefinition<CodeReviewDefectsFieldName, CodeReviewDefectsViewField> = {
	title: LIST_TITLE,
	description: "Code review defects list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionCodeReviewDefects(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionCodeReviewDefects;
