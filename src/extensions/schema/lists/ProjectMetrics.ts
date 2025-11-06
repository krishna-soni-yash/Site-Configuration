import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    RequiredListsProvision,
    ensureListProvision,
    ListProvisionDefinition,
    FieldDefinition
} from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.ProjectMetrics;

const projectMetricsFieldNames = [
    "VersionId",
    "Applicability",
    "AssociatedPPM",
    "BaselineAndRevisionFrequency",
    "BG",
    "CausalAnalysisTrigger",
    "DataCollectionFrequency",
    "DataAnalysisFrequency",
    "DataInput",
    "DataSource",
    "Goal",
    "HasSubProcess",
    "InterpretationGuidelines",
    "IsActive",
    "LSL",
    "Metrics",
    "MetricsFormulae",
    "OrgCausalAnalysisTrigger",
    "OrgInterpretationGuidelines",
    "OrgStatistical",
    "PerformanceGoals",
    "PG",
    "Priority",
    "ProbabilityOfSuccessThreshold",
    "Process",
    "ProjectType",
    "Quantitative",
    "Statistical",
    "SubApplicability",
    "SubBaselineAndRevisionFrequency",
    "SubDataAnalysisFrequency",
    "SubDataCollectionFrequency",
    "SubDataInput",
    "SubDataSource",
    "SubGoal",
    "SubLSL",
    "SubMetrics",
    "SubMetricsFormulae",
    "Subprocess",
    "SubUnitOfMeasure",
    "SubUSL",
    "UnitOfMeasure",
    "USL"
] as const;

type ProjectMetricsFieldName = typeof projectMetricsFieldNames[number];
type ProjectMetricsViewField = ProjectMetricsFieldName | "LinkTitle" | "Modified" | "Created";

const yesNoChoiceFieldNames: readonly ProjectMetricsFieldName[] = [
    "Applicability",
    "HasSubProcess"
] as const;

const textFieldNames: readonly ProjectMetricsFieldName[] = [
    "AssociatedPPM",
    "BaselineAndRevisionFrequency",
    "BG",
    "DataCollectionFrequency",
    "DataAnalysisFrequency",
    "DataSource",
    "Goal",
    "OrgStatistical",
    "PerformanceGoals",
    "PG",
    "Priority",
    "ProbabilityOfSuccessThreshold",
    "Process",
    "ProjectType",
    "Quantitative",
    "Statistical",
    "SubApplicability",
    "SubBaselineAndRevisionFrequency",
    "SubDataAnalysisFrequency",
    "SubDataCollectionFrequency",
    "SubDataSource",
    "SubGoal",
    "SubMetrics",
    "Subprocess",
    "SubUnitOfMeasure",
    "UnitOfMeasure"
] as const;

const noteFieldNames: readonly ProjectMetricsFieldName[] = [
    "CausalAnalysisTrigger",
    "DataInput",
    "InterpretationGuidelines",
    "MetricsFormulae",
    "OrgCausalAnalysisTrigger",
    "OrgInterpretationGuidelines",
    "SubDataInput",
    "SubMetricsFormulae"
] as const;

const numberFieldNames: readonly ProjectMetricsFieldName[] = [
    "LSL",
    "SubLSL",
    "SubUSL",
    "USL"
] as const;

const booleanFieldNames: readonly ProjectMetricsFieldName[] = ["IsActive"] as const;

const defaultViewFields: readonly ProjectMetricsViewField[] = [
    "LinkTitle",
    "Applicability",
    "DataSource",
    "AssociatedPPM",
    "BaselineAndRevisionFrequency",
    "BG",
    "CausalAnalysisTrigger",
    "DataAnalysisFrequency",
    "DataCollectionFrequency",
    "Goal",
    "HasSubProcess",
    "InterpretationGuidelines",
    "IsActive",
    "USL",
    "LSL",
    "Metrics",
    "MetricsFormulae",
    "OrgCausalAnalysisTrigger",
    "OrgInterpretationGuidelines",
    "OrgStatistical",
    "PerformanceGoals",
    "PG",
    "Priority",
    "ProbabilityOfSuccessThreshold",
    "Process",
    "ProjectType",
    "Quantitative",
    "UnitOfMeasure",
    "Statistical",
    "SubApplicability",
    "SubBaselineAndRevisionFrequency",
    "SubDataAnalysisFrequency",
    "SubDataCollectionFrequency",
    "SubDataInput",
    "SubDataSource",
    "SubGoal",
    "SubLSL",
    "SubMetrics",
    "SubMetricsFormulae",
    "Subprocess",
    "SubUnitOfMeasure",
    "SubUSL",
    "DataInput",
    "VersionId",
    "Modified",
    "Created"
] as const;

const yesNoChoices = ["Yes", "No"] as const;
const yesNoChoicesXml = yesNoChoices.map(choice => `<CHOICE>${choice}</CHOICE>`).join("");

function buildFieldDefinitions(logsListId: string): FieldDefinition<ProjectMetricsFieldName>[] {
    const definitions: FieldDefinition<ProjectMetricsFieldName>[] = [
        {
            internalName: "VersionId",
            schemaXml: `<Field Type='Lookup' Name='VersionId' StaticName='VersionId' DisplayName='VersionId' List='${logsListId}' ShowField='ID' />`
        }
    ];

    for (const internalName of yesNoChoiceFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Choice' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' Format='Dropdown'><CHOICES>${yesNoChoicesXml}</CHOICES></Field>`
        });
    }

    for (const internalName of textFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Text' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' MaxLength='255' />`
        });
    }

    for (const internalName of noteFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Note' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' NumLines='6' RichText='FALSE' />`
        });
    }

    for (const internalName of numberFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Number' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' />`
        });
    }

    for (const internalName of booleanFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Boolean' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' />`
        });
    }

    return definitions;
}

export async function provisionProjectMetrics(sp: SPFI): Promise<void> {
    let logsListId: string | undefined;

    try {
        const logsListInfo = await sp.web.lists.getByTitle(RequiredListsProvision.ProjectMetricLogs).select("Id")();
        logsListId = `${logsListInfo.Id}`;
    } catch (error) {
        const ensureResult = await sp.web.lists.ensure(
            RequiredListsProvision.ProjectMetricLogs,
            "Project metrics logs list",
            100
        );
        const ensuredInfo = await ensureResult.list.select("Id")();
        logsListId = `${ensuredInfo.Id}`;
    }

    if (!logsListId) {
        throw new Error("Unable to resolve Project Metric Logs list identifier.");
    }

    if (!logsListId.startsWith("{")) {
        logsListId = `{${logsListId}}`;
    }

    const fields = buildFieldDefinitions(logsListId);

    const definition: ListProvisionDefinition<ProjectMetricsFieldName, ProjectMetricsViewField> = {
        title: LIST_TITLE,
        description: "Project metrics list",
        templateId: 100,
        fields,
        defaultViewFields
    };

    await ensureListProvision(sp, definition);
}

export default provisionProjectMetrics;