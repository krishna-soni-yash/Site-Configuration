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

const LIST_TITLE = RequiredListsProvision.TestingDefects;

type TestingDefectsFieldName =
	| "Test Scenario ID"
	| "Test Case ID"
	| "Defect Description"
	| "Review method"
	| "Defect detected on"
	| "Defect Detected by"
	| "Defect Status"
	| "Defect Type"
	| "Injected Phase"
	| "Identified Phase"
	| "Defect Classification"
	| "Severity"
	| "Priority"
	| "Root Cause"
	| "Defect Closure Date"
	| "Remarks"
	| "Defect Fixed By";

type TestingDefectsViewField = TestingDefectsFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<TestingDefectsFieldName>[] = [
	{
		internalName: "Test Scenario ID",
		schemaXml: `<Field Type='Text' Name='Test Scenario ID' StaticName='Test Scenario ID' DisplayName='Test Scenario ID' MaxLength='255' />`
	},
	{
		internalName: "Test Case ID",
		schemaXml: `<Field Type='Text' Name='Test Case ID' StaticName='Test Case ID' DisplayName='Test Case ID' MaxLength='255' />`
	},
	{
		internalName: "Defect Description",
		schemaXml: `<Field Type='Note' Name='Defect Description' StaticName='Defect Description' DisplayName='Defect Description' NumLines='6' RichText='FALSE' />`
	},
	{
		internalName: "Review method",
		schemaXml: `<Field Type='Choice' Name='Review method' StaticName='Review method' DisplayName='Testing Type' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Unit Testing</CHOICE><CHOICE>Integration Testing</CHOICE><CHOICE>System Testing</CHOICE><CHOICE>Regression Testing</CHOICE><CHOICE>User Acceptance Testing</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Defect detected on",
		schemaXml: `<Field Type='DateTime' Name='Defect detected on' StaticName='Defect detected on' DisplayName='Defect detected on' Format='DateOnly' />`
	},
	{
		internalName: "Defect Detected by",
		schemaXml: `<Field Type='User' Name='Defect Detected by' StaticName='Defect Detected by' DisplayName='Defect Detected by' UserSelectionMode='PeopleOnly' />`
	},
	{
		internalName: "Defect Status",
		schemaXml: `<Field Type='Choice' Name='Defect Status' StaticName='Defect Status' DisplayName='Defect Status' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>New</CHOICE><CHOICE>Open</CHOICE><CHOICE>Resolved- Fixed</CHOICE><CHOICE>Resolved- Invalid</CHOICE><CHOICE>Resolved- Won&apos;t Fix</CHOICE><CHOICE>Resolved- Duplicate</CHOICE><CHOICE>Resolved- Moved</CHOICE><CHOICE>Retest</CHOICE><CHOICE>Reopen</CHOICE><CHOICE>On Hold</CHOICE><CHOICE>Deferred</CHOICE><CHOICE>Closed</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Defect Type",
		schemaXml: `<Field Type='Choice' Name='Defect Type' StaticName='Defect Type' DisplayName='Defect Type' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Requirement</CHOICE><CHOICE>Design</CHOICE><CHOICE>Function</CHOICE><CHOICE>Interface</CHOICE><CHOICE>Checking</CHOICE><CHOICE>Algorithm</CHOICE><CHOICE>Assignment</CHOICE><CHOICE>Timing/Serialization</CHOICE><CHOICE>Build</CHOICE><CHOICE>Documentation</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Injected Phase",
		schemaXml: `<Field Type='Choice' Name='Injected Phase' StaticName='Injected Phase' DisplayName='Defect Origin Phase' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Requirement</CHOICE><CHOICE>Design</CHOICE><CHOICE>Code</CHOICE><CHOICE>Unit Test</CHOICE><CHOICE>Integration Test</CHOICE><CHOICE>System Test</CHOICE><CHOICE>Acceptance Test</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Identified Phase",
		schemaXml: `<Field Type='Choice' Name='Identified Phase' StaticName='Identified Phase' DisplayName='Defect Detected Phase' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Requirement</CHOICE><CHOICE>Design</CHOICE><CHOICE>Code</CHOICE><CHOICE>Unit Test</CHOICE><CHOICE>Integration Test</CHOICE><CHOICE>System Test</CHOICE><CHOICE>Regression Test</CHOICE><CHOICE>Acceptance Test</CHOICE><CHOICE>Post Production</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Defect Classification",
		schemaXml: `<Field Type='Choice' Name='Defect Classification' StaticName='Defect Classification' DisplayName='Defect Classification' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Requirements - Missing requirement</CHOICE><CHOICE>Requirements - Incomplete requirement</CHOICE><CHOICE>Requirements - Incorrect requirement</CHOICE><CHOICE>Requirements - Not meeting the standard</CHOICE><CHOICE>Design - Data Model (ER Model)/Database Design</CHOICE><CHOICE>Design - Security Design</CHOICE><CHOICE>Design - Compatibility</CHOICE><CHOICE>Design - Third Party Components(COTS) related design</CHOICE><CHOICE>Design - User Documentation/Help Design</CHOICE><CHOICE>Design - Test Cases/Test Scripts</CHOICE><CHOICE>Design issues</CHOICE><CHOICE>Data Validation</CHOICE><CHOICE>Logical</CHOICE><CHOICE>Computational</CHOICE><CHOICE>User Interface Presentation</CHOICE><CHOICE>Input/output</CHOICE><CHOICE>Error Handling/Exception Handling</CHOICE><CHOICE>Initialization</CHOICE><CHOICE>Installation/Configuration</CHOICE><CHOICE>Performance/Load/Stress</CHOICE><CHOICE>Online help / Error messages</CHOICE><CHOICE>Third Party Software</CHOICE><CHOICE>Not Meeting Coding Standard</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Severity",
		schemaXml: `<Field Type='Choice' Name='Severity' StaticName='Severity' DisplayName='Severity' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Critical</CHOICE><CHOICE>Major</CHOICE><CHOICE>Medium</CHOICE><CHOICE>Minor</CHOICE><CHOICE>Enhancements/Suggestions</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Priority",
		schemaXml: `<Field Type='Choice' Name='Priority' StaticName='Priority' DisplayName='Priority' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Urgent</CHOICE><CHOICE>High</CHOICE><CHOICE>Medium</CHOICE><CHOICE>Low</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Root Cause",
		schemaXml: `<Field Type='Choice' Name='Root Cause' StaticName='Root Cause' DisplayName='Root Cause' Format='Dropdown'><CHOICES><CHOICE>--Select--</CHOICE><CHOICE>Oversight</CHOICE><CHOICE>Understanding Error</CHOICE><CHOICE>Communication Gap</CHOICE><CHOICE>External Factor</CHOICE><CHOICE>Inadequate Knowledge</CHOICE><CHOICE>Inadequate Skills</CHOICE></CHOICES></Field>`
	},
	{
		internalName: "Defect Closure Date",
		schemaXml: `<Field Type='DateTime' Name='Defect Closure Date' StaticName='Defect Closure Date' DisplayName='Defect Closure Date' Format='DateOnly' />`
	},
	{
		internalName: "Remarks",
		schemaXml: `<Field Type='Text' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' MaxLength='255' />`
	},
	{
		internalName: "Defect Fixed By",
		schemaXml: `<Field Type='User' Name='Defect Fixed By' StaticName='Defect Fixed By' DisplayName='Defect Fixed By' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
	}
] as const;

const defaultViewFields: readonly TestingDefectsViewField[] = [
	"LinkTitle",
	"Test Scenario ID",
	"Test Case ID",
	"Defect Description",
	"Review method",
	"Defect detected on",
	"Defect Detected by",
	"Defect Status",
	"Defect Type",
	"Injected Phase",
	"Identified Phase",
	"Defect Classification",
	"Severity",
	"Priority",
	"Root Cause",
	"Defect Closure Date",
	"Remarks",
	"Defect Fixed By"
] as const;

const definition: ListProvisionDefinition<TestingDefectsFieldName, TestingDefectsViewField> = {
	title: LIST_TITLE,
	description: "Testing defects list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionTestingDefects(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	await list.fields.getByInternalNameOrTitle("Title").update({ Title: "Requirement #" });
}

export default provisionTestingDefects;
