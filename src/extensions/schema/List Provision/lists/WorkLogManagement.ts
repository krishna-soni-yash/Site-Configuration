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

const LIST_TITLE = RequiredListsProvision.WorkLogManagement;

type WorkLogFieldName =
    | "ReqTitle"
    | "AssignedToUsers"
    | "ProjectType"
    | "WorkItemNo"
    | "Simple"
    | "Medium"
    | "Complex"
    | "VeryComplex"
    | "ComplexityPoints"
    | "AdjustedComplexityPoint"
    | "AppAndEnvAdjustmentFactor"
    | "SkillAdjustmentFactor"
    | "ReusabilityOfDesignAndCode"
    | "ExtentOfAutomation"
    | "AdjustedEffort"
    | "BaseEffort"
    | "CalculatedPlannedEffort"
    | "ActualPlannedEffort"
    | "PlannedStartDate"
    | "PlannedEndDate"
    | "Status"
    | "Remarks";

type WorkLogViewField = WorkLogFieldName | "ID" | "Modified" | "Editor";

const fieldDefinitions: readonly FieldDefinition<WorkLogFieldName>[] = [
    { internalName: "ReqTitle", schemaXml: `<Field Type='Text' Name='ReqTitle' StaticName='ReqTitle' DisplayName='ReqTitle' MaxLength='255' />` },
    { internalName: "AssignedToUsers", schemaXml: `<Field Type='User' Name='AssignedToUsers' StaticName='AssignedToUsers' DisplayName='AssignedToUsers' UserSelectionMode='PeopleOnly' Mult='TRUE' />` },
    { internalName: "ProjectType", schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />` },
    { internalName: "WorkItemNo", schemaXml: `<Field Type='Text' Name='WorkItemNo' StaticName='WorkItemNo' DisplayName='WorkItemNo' MaxLength='255' />` },
    { internalName: "Simple", schemaXml: `<Field Type='Number' Name='Simple' StaticName='Simple' DisplayName='Simple' Decimals='2' />` },
    { internalName: "Medium", schemaXml: `<Field Type='Number' Name='Medium' StaticName='Medium' DisplayName='Medium' Decimals='2' />` },
    { internalName: "Complex", schemaXml: `<Field Type='Number' Name='Complex' StaticName='Complex' DisplayName='Complex' Decimals='2' />` },
    { internalName: "VeryComplex", schemaXml: `<Field Type='Number' Name='VeryComplex' StaticName='VeryComplex' DisplayName='VeryComplex' Decimals='2' />` },
    { internalName: "ComplexityPoints", schemaXml: `<Field Type='Number' Name='ComplexityPoints' StaticName='ComplexityPoints' DisplayName='ComplexityPoints' Decimals='2' />` },
    { internalName: "AdjustedComplexityPoint", schemaXml: `<Field Type='Number' Name='AdjustedComplexityPoint' StaticName='AdjustedComplexityPoint' DisplayName='AdjustedComplexityPoint' Decimals='2' />` },
    { internalName: "AppAndEnvAdjustmentFactor", schemaXml: `<Field Type='Number' Name='AppAndEnvAdjustmentFactor' StaticName='AppAndEnvAdjustmentFactor' DisplayName='AppAndEnvAdjustmentFactor' Decimals='2' />` },
    { internalName: "SkillAdjustmentFactor", schemaXml: `<Field Type='Number' Name='SkillAdjustmentFactor' StaticName='SkillAdjustmentFactor' DisplayName='SkillAdjustmentFactor' Decimals='2' />` },
    { internalName: "ReusabilityOfDesignAndCode", schemaXml: `<Field Type='Number' Name='ReusabilityOfDesignAndCode' StaticName='ReusabilityOfDesignAndCode' DisplayName='ReusabilityOfDesignAndCode' Decimals='2' />` },
    { internalName: "ExtentOfAutomation", schemaXml: `<Field Type='Number' Name='ExtentOfAutomation' StaticName='ExtentOfAutomation' DisplayName='ExtentOfAutomation' Decimals='2' />` },
    { internalName: "AdjustedEffort", schemaXml: `<Field Type='Number' Name='AdjustedEffort' StaticName='AdjustedEffort' DisplayName='AdjustedEffort' Decimals='2' />` },
    { internalName: "BaseEffort", schemaXml: `<Field Type='Number' Name='BaseEffort' StaticName='BaseEffort' DisplayName='BaseEffort' Decimals='2' />` },
    { internalName: "CalculatedPlannedEffort", schemaXml: `<Field Type='Number' Name='CalculatedPlannedEffort' StaticName='CalculatedPlannedEffort' DisplayName='CalculatedPlannedEffort' Decimals='2' />` },
    { internalName: "ActualPlannedEffort", schemaXml: `<Field Type='Number' Name='ActualPlannedEffort' StaticName='ActualPlannedEffort' DisplayName='ActualPlannedEffort' Decimals='2' />` },
    { internalName: "PlannedStartDate", schemaXml: `<Field Type='DateTime' Name='PlannedStartDate' StaticName='PlannedStartDate' DisplayName='PlannedStartDate' Format='DateOnly' />` },
    { internalName: "PlannedEndDate", schemaXml: `<Field Type='DateTime' Name='PlannedEndDate' StaticName='PlannedEndDate' DisplayName='PlannedEndDate' Format='DateOnly' />` },
    { internalName: "Status", schemaXml: `<Field Type='Text' Name='Status' StaticName='Status' DisplayName='Status' MaxLength='255' />` },
    { internalName: "Remarks", schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />` }
] as const;

const defaultViewFields: readonly WorkLogViewField[] = [
    "ID",
    "ReqTitle",
    "AssignedToUsers",
    "ProjectType",
    "WorkItemNo",
    "Simple",
    "Medium",
    "Complex",
    "VeryComplex",
    "ComplexityPoints",
    "AdjustedComplexityPoint",
    "AppAndEnvAdjustmentFactor",
    "SkillAdjustmentFactor",
    "ReusabilityOfDesignAndCode",
    "ExtentOfAutomation",
    "AdjustedEffort",
    "BaseEffort",
    "CalculatedPlannedEffort",
    "ActualPlannedEffort",
    "PlannedStartDate",
    "PlannedEndDate",
    "Status",
    "Remarks",
    "Modified",
    "Editor"
] as const;

const definition: ListProvisionDefinition<WorkLogFieldName, WorkLogViewField> = {
    title: LIST_TITLE,
    description: "Work log management",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionWorkLogManagement(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionWorkLogManagement;
