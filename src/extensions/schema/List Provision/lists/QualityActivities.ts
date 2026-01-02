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

const LIST_TITLE = RequiredListsProvision.QualityActivities;

type QualityActivitiesFieldName =
    | "Task"
    | "PlannedDate"
    | "ActualDate"
    | "Status"
    | "RiskSummaryCol1"
    | "RiskSummaryCol2"
    | "RiskSummaryCol3"
    | "ProcessGoal"
    | "SubProcessGoal"
    | "CSI"
    | "ResourceUtilization"
    | "COQ"
    | "OpenNCs"
    | "OpenObservations"
    | "OpenFcFindings"
    | "PCI"
    | "Remarks";

type QualityActivitiesViewField = QualityActivitiesFieldName | "ID" | "Modified" | "Editor";

const taskChoices = ["Facilitation", "PMWB Review", "Project Meeting", "Internal Audit", "Training", "BUH Connect", "External Audit"] as const;
const statusChoices = ["Open", "Closed"] as const;
const goalChoices = ["Not Met", "Met"] as const;

function buildChoiceFieldSchema(name: string, displayName: string, choices: readonly string[]): string {
    const choicesXml = choices.map((c) => `<CHOICE>${c}</CHOICE>`).join("");
    return `<Field Type='Choice' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='Dropdown'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

const fieldDefinitions: readonly FieldDefinition<QualityActivitiesFieldName>[] = [
    { internalName: "Task", schemaXml: buildChoiceFieldSchema("Task", "Task", taskChoices as readonly string[]) },
    { internalName: "PlannedDate", schemaXml: `<Field Type='DateTime' Name='PlannedDate' StaticName='PlannedDate' DisplayName='PlannedDate' Format='DateOnly' />` },
    { internalName: "ActualDate", schemaXml: `<Field Type='DateTime' Name='ActualDate' StaticName='ActualDate' DisplayName='ActualDate' Format='DateOnly' />` },
    { internalName: "Status", schemaXml: buildChoiceFieldSchema("Status", "Status", statusChoices as readonly string[]) },
    { internalName: "RiskSummaryCol1", schemaXml: `<Field Type='Number' Name='RiskSummaryCol1' StaticName='RiskSummaryCol1' DisplayName='Risk Summary (RE >= 80)' Decimals='2' />` },
    { internalName: "RiskSummaryCol2", schemaXml: `<Field Type='Number' Name='RiskSummaryCol2' StaticName='RiskSummaryCol2' DisplayName='Risk Summary (RE >= 60 & < 80)' Decimals='2' />` },
    { internalName: "RiskSummaryCol3", schemaXml: `<Field Type='Number' Name='RiskSummaryCol3' StaticName='RiskSummaryCol3' DisplayName='Risk Summary (RE >= 0 & < 60)' Decimals='2' />` },
    { internalName: "ProcessGoal", schemaXml: buildChoiceFieldSchema("ProcessGoal", "ProcessGoal", goalChoices as readonly string[]) },
    { internalName: "SubProcessGoal", schemaXml: buildChoiceFieldSchema("SubProcessGoal", "SubProcessGoal", goalChoices as readonly string[]) },
    { internalName: "CSI", schemaXml: `<Field Type='Text' Name='CSI' StaticName='CSI' DisplayName='CSI' MaxLength='255' />` },
    { internalName: "ResourceUtilization", schemaXml: `<Field Type='Number' Name='ResourceUtilization' StaticName='ResourceUtilization' DisplayName='Resource Utilization(%)' Decimals='2' />` },
    { internalName: "COQ", schemaXml: `<Field Type='Number' Name='COQ' StaticName='COQ' DisplayName='COQ(%)' Decimals='2' />` },
    { internalName: "OpenNCs", schemaXml: `<Field Type='Number' Name='OpenNCs' StaticName='OpenNCs' DisplayName='No of Open NCs' Decimals='0' />` },
    { internalName: "OpenObservations", schemaXml: `<Field Type='Number' Name='OpenObservations' StaticName='OpenObservations' DisplayName='No of Open Observations' Decimals='0' />` },
    { internalName: "OpenFcFindings", schemaXml: `<Field Type='Number' Name='OpenFcFindings' StaticName='OpenFcFindings' DisplayName='No of Open Facilitation Findings' Decimals='0' />` },
    { internalName: "PCI", schemaXml: `<Field Type='Number' Name='PCI' StaticName='PCI' DisplayName='Process Compliance Index (PCI)' Decimals='2' />` },
    { internalName: "Remarks", schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />` }
] as const;

const defaultViewFields: readonly QualityActivitiesViewField[] = [
    "ID",
    "Task",
    "PlannedDate",
    "ActualDate",
    "Status",
    "RiskSummaryCol1",
    "RiskSummaryCol2",
    "RiskSummaryCol3",
    "ProcessGoal",
    "SubProcessGoal",
    "CSI",
    "ResourceUtilization",
    "COQ",
    "OpenNCs",
    "OpenObservations",
    "OpenFcFindings",
    "PCI",
    "Remarks",
    "Modified",
    "Editor"
] as const;

const definition: ListProvisionDefinition<QualityActivitiesFieldName, QualityActivitiesViewField> = {
    title: LIST_TITLE,
    description: "Quality activities",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionQualityActivities(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionQualityActivities;
