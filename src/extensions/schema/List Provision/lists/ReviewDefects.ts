/*eslint-disable*/
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/folders";
import {
	ensureListProvision,
	FieldDefinition,
	ListProvisionDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.ReviewDefects;

type ReviewDefectsFieldName =
	| "VersionNo"
	| "TypeOfReview"
	| "TypeOfReviewArtifact"
	| "ReviewCompletionDate"
	| "PlannedClosureDate"
	| "ActualClosureDate"
	| "Reviewers"
	| "DefectDescription"
	| "Category"
	| "DefectType"
	| "DefectClassification"
	| "DefectOriginPhase"
	| "DefectDetectedPhase"
	| "Severity"
	| "Status"
	| "ReviewResults"
	| "CorrectionCorrectiveAction"
	| "Remarks";

type ReviewDefectsViewField = ReviewDefectsFieldName | "ID" | "Modified" | "Editor" | "Title";

const typeOfReviewChoices = ["Walkthrough", "Inspection"] as const;
const typeOfReviewArtifactChoices = [
	"PIN",
	"PDP",
	"PRR",
	"PMP",
	"MPP",
	"PIP",
	"SRS",
	"HLD",
	"LLD",
	"Test Cases",
	"PM-workbook",
	"FS",
	"TS",
	"Unit Test cases",
	"Test Plan",
	"Others"
] as const;
const categoryChoices = ["--Select--", "Observation", "Suggestion", "Improvement", "Defect"] as const;
const defectTypeChoices = ["--Select--", "Requirements", "Design", "Others"] as const;
const defectClassificationChoices = [
	"--Select--",
	"Requirements - Incomplete requirement",
	"Requirements - Incorrect requirement",
	"Design - Data Model (ER Model)/Database Design",
	"Design - Security Design",
	"Design - Compatibility",
	"Design - Third Party Components(COTS) related design",
	"Design - User Documentation/Help Design",
	"Design - Test Cases/Test Scripts",
	"Project Planning",
	"Project Tailoring",
	"Resource Planning",
	"Project Initiation",
	"Configuration Issues",
	"Work Plan",
	"Estimation",
	"Test Plan"
] as const;
const severityChoices = ["--Select--", "Critical", "Moderate", "Minor"] as const;
const statusChoices = ["--Select--", "Open", "Closed", "Deferred", "In-Process"] as const;
const reviewResultsChoices = ["--Select--", "Open", "Closed", "Deferred", "In-Process"] as const;
const defectPhaseChoices = [
	"--Select--",
	"Requirements Specification",
	"Designs",
	"Code",
	"Unit Testing",
	"Integration Testing",
	"System Testing",
	"Acceptance Testing",
	"Others"
] as const;

function buildChoiceFieldSchema(name: string, displayName: string, choices: readonly string[]): string {
	const choicesXml = choices.map((choice) => `<CHOICE>${choice}</CHOICE>`).join("");
	return `<Field Type='Choice' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='Dropdown'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

function buildFieldDefinitions(): FieldDefinition<ReviewDefectsFieldName>[] {
	return [
		{
			internalName: "VersionNo",
			schemaXml: `<Field Type='Text' Name='VersionNo' StaticName='VersionNo' DisplayName='Version No' MaxLength='255' />`
		},
		{
			internalName: "TypeOfReview",
			schemaXml: buildChoiceFieldSchema("TypeOfReview", "Type of Review", typeOfReviewChoices as readonly string[])
		},
		{
			internalName: "TypeOfReviewArtifact",
			schemaXml: buildChoiceFieldSchema("TypeOfReviewArtifact", "Type of Review Artifact", typeOfReviewArtifactChoices as readonly string[])
		},
		{
			internalName: "ReviewCompletionDate",
			schemaXml: `<Field Type='DateTime' Name='ReviewCompletionDate' StaticName='ReviewCompletionDate' DisplayName='Review Completion Date' Format='DateOnly' />`
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
			internalName: "Reviewers",
			schemaXml: `<Field Type='User' Name='Reviewers' StaticName='Reviewers' DisplayName='Reviewers' UserSelectionMode='PeopleOnly' Mult='TRUE' />`
		},
		{
			internalName: "DefectDescription",
			schemaXml: `<Field Type='Note' Name='DefectDescription' StaticName='DefectDescription' DisplayName='Defect Description' NumLines='6' RichText='FALSE' />`
		},
		{
			internalName: "Category",
			schemaXml: buildChoiceFieldSchema("Category", "Category", categoryChoices as readonly string[])
		},
		{
			internalName: "DefectType",
			schemaXml: buildChoiceFieldSchema("DefectType", "Defect Type", defectTypeChoices as readonly string[])
		},
		{
			internalName: "DefectClassification",
			schemaXml: buildChoiceFieldSchema("DefectClassification", "Defect Classification", defectClassificationChoices as readonly string[])
		},
		{
			internalName: "DefectOriginPhase",
			schemaXml: buildChoiceFieldSchema("DefectOriginPhase", "Defect Origin Phase", defectPhaseChoices as readonly string[])
		},
		{
			internalName: "DefectDetectedPhase",
			schemaXml: buildChoiceFieldSchema("DefectDetectedPhase", "Defect Detected Phase", defectPhaseChoices as readonly string[])
		},
		{
			internalName: "Severity",
			schemaXml: buildChoiceFieldSchema("Severity", "Severity", severityChoices as readonly string[])
		},
		{
			internalName: "Status",
			schemaXml: buildChoiceFieldSchema("Status", "Status", statusChoices as readonly string[])
		},
		{
			internalName: "ReviewResults",
			schemaXml: buildChoiceFieldSchema("ReviewResults", "Review Results", reviewResultsChoices as readonly string[])
		},
		{
			internalName: "CorrectionCorrectiveAction",
			schemaXml: `<Field Type='Note' Name='CorrectionCorrectiveAction' StaticName='CorrectionCorrectiveAction' DisplayName='Correction / Corrective Action' NumLines='6' RichText='FALSE' />`
		},
		{
			internalName: "Remarks",
			schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
		}
	];
}

const defaultViewFields: readonly ReviewDefectsViewField[] = [
	"Title",
	"VersionNo",
	"TypeOfReview",
	"TypeOfReviewArtifact",
	"ReviewCompletionDate",
	"PlannedClosureDate",
	"ActualClosureDate",
	"Reviewers",
	"DefectDescription",
	"Category",
	"DefectType",
	"DefectClassification",
	"DefectOriginPhase",
	"DefectDetectedPhase",
	"Severity",
	"Status",
	"ReviewResults",
	"CorrectionCorrectiveAction",
	"Remarks"
] as const;

async function ensureArtifactNameField(sp: SPFI): Promise<void> {
	try {
		const list = sp.web.lists.getByTitle(LIST_TITLE);
		await list.fields.getByInternalNameOrTitle("Title").update({ Title: "Artifact Name" });
	} catch (error) {
		console.warn(`Failed to rename Title field for list ${LIST_TITLE}:`, error);
	}
}

export async function provisionReviewDefects(sp: SPFI): Promise<void> {
	const fields = buildFieldDefinitions();

	const definition: ListProvisionDefinition<ReviewDefectsFieldName, ReviewDefectsViewField> = {
		title: LIST_TITLE,
		description: "Review defects log",
		templateId: 100,
		fields,
		defaultViewFields
	};

	await ensureListProvision(sp, definition);
	await ensureArtifactNameField(sp);
}

export default provisionReviewDefects;
