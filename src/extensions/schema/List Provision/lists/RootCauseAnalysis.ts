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

const LIST_TITLE = RequiredListsProvision.RootCauseAnalysis;

type RootCauseAnalysisFieldName =
    | "ProblemStatementNumber"
    | "CauseCategory"
    | "Cause"
    | "RCASource"
    | "RCAPriority"
    | "RootCause"
    | "RCATechniqueUsedAndReference"
    | "PerformanceBeforeActionPlan"
    | "PerformanceAfterActionPlan"
    | "RelatedMetric"
    | "RCATypeOfAction"
    | "ActionPlanCorrective"
    | "ActionPlanPreventive"
    | "ActionPlanCorrection"
    | "ResponsibilityCorrection"
    | "ResponsibilityCorrective"
    | "ResponsibilityPreventive"
    | "PlannedClosureDateCorrection"
    | "ActualClosureDateCorrective"
    | "PlannedClosureDatePreventive"
    | "ActualClosureDatePreventive"
    | "RelatedSubMetric"
    | "ActualClosureDateCorrection"
    | "Remarks"
    | "TypeOfAction";
    
type RootCauseAnalysisViewField = RootCauseAnalysisFieldName | "ID" | "Modified" | "Editor" | "Attachments";

const causeCategoryChoices = ["Special", "Common"] as const;
const rcaSourceChoices = [
    "Audit Findings",
    "Metrics",
    "Review Findings",
    "Testing Findings",
    "CustomerFeedback"
] as const;
const rcaPriorityChoices = ["High", "Medium", "Low"] as const;
const rcaTypeOfActionChoices = ["Correction", "Corrective Action", "Preventive Action"] as const;
const typeOfActionChoices = ["Mitigation", "Contingency"] as const;

function buildChoiceFieldSchema(name: string, displayName: string, choices: readonly string[]): string {
    const choicesXml = choices.map((choice) => `<CHOICE>${choice}</CHOICE>`).join("");
    return `<Field Type='Choice' Name='${name}' StaticName='${name}' DisplayName='${displayName}' Format='Dropdown'><CHOICES>${choicesXml}</CHOICES></Field>`;
}

const fieldDefinitions: readonly FieldDefinition<RootCauseAnalysisFieldName>[] = [
    {
        internalName: "ProblemStatementNumber",
        schemaXml: `<Field Type='Text' Name='ProblemStatementNumber' StaticName='ProblemStatementNumber' DisplayName='ProblemStatementNumber' MaxLength='255' />`
    },
    {
        internalName: "CauseCategory",
        schemaXml: buildChoiceFieldSchema("CauseCategory", "CauseCategory", causeCategoryChoices as readonly string[])
    },
    {
        internalName: "Cause",
        schemaXml: `<Field Type='Note' Name='Cause' StaticName='Cause' DisplayName='Cause' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "RCASource",
        schemaXml: buildChoiceFieldSchema("RCASource", "RCASource", rcaSourceChoices as readonly string[])
    },
    {
        internalName: "RCAPriority",
        schemaXml: buildChoiceFieldSchema("RCAPriority", "RCAPriority", rcaPriorityChoices as readonly string[])
    },
    {
        internalName: "RootCause",
        schemaXml: `<Field Type='Note' Name='RootCause' StaticName='RootCause' DisplayName='RootCause' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "RCATechniqueUsedAndReference",
        schemaXml: `<Field Type='Note' Name='RCATechniqueUsedAndReference' StaticName='RCATechniqueUsedAndReference' DisplayName='RCATechniqueUsedAndReference' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "PerformanceBeforeActionPlan",
        schemaXml: `<Field Type='Note' Name='PerformanceBeforeActionPlan' StaticName='PerformanceBeforeActionPlan' DisplayName='PerformanceBeforeActionPlan' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "PerformanceAfterActionPlan",
        schemaXml: `<Field Type='Note' Name='PerformanceAfterActionPlan' StaticName='PerformanceAfterActionPlan' DisplayName='PerformanceAfterActionPlan' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "RelatedMetric",
        schemaXml: `<Field Type='Text' Name='RelatedMetric' StaticName='RelatedMetric' DisplayName='RelatedMetric' MaxLength='255' />`
    },
    {
        internalName: "RCATypeOfAction",
        schemaXml: buildChoiceFieldSchema("RCATypeOfAction", "RCATypeOfAction", rcaTypeOfActionChoices as readonly string[])
    },
    {
        internalName: "ActionPlanCorrective",
        schemaXml: `<Field Type='Note' Name='ActionPlanCorrective' StaticName='ActionPlanCorrective' DisplayName='ActionPlanCorrective' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "ActionPlanPreventive",
        schemaXml: `<Field Type='Note' Name='ActionPlanPreventive' StaticName='ActionPlanPreventive' DisplayName='ActionPlanPreventive' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "ActionPlanCorrection",
        schemaXml: `<Field Type='Note' Name='ActionPlanCorrection' StaticName='ActionPlanCorrection' DisplayName='ActionPlanCorrection' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "ResponsibilityCorrection",
        schemaXml: `<Field Type='User' Name='ResponsibilityCorrection' StaticName='ResponsibilityCorrection' DisplayName='ResponsibilityCorrection' UserSelectionMode='PeopleOnly' Mult='FALSE' />`
    },
    {
        internalName: "ResponsibilityCorrective",
        schemaXml: `<Field Type='User' Name='ResponsibilityCorrective' StaticName='ResponsibilityCorrective' DisplayName='ResponsibilityCorrective' UserSelectionMode='PeopleOnly' Mult='FALSE' />`
    },
    {
        internalName: "ResponsibilityPreventive",
        schemaXml: `<Field Type='User' Name='ResponsibilityPreventive' StaticName='ResponsibilityPreventive' DisplayName='ResponsibilityPreventive' UserSelectionMode='PeopleOnly' Mult='FALSE' />`
    },
    {
        internalName: "PlannedClosureDateCorrection",
        schemaXml: `<Field Type='DateTime' Name='PlannedClosureDateCorrection' StaticName='PlannedClosureDateCorrection' DisplayName='PlannedClosureDateCorrection' Format='DateOnly' />`
    },
    {
        internalName: "ActualClosureDateCorrective",
        schemaXml: `<Field Type='DateTime' Name='ActualClosureDateCorrective' StaticName='ActualClosureDateCorrective' DisplayName='ActualClosureDateCorrective' Format='DateOnly' />`
    }
    ,
    {
        internalName: "ActualClosureDateCorrection",
        schemaXml: `<Field Type='DateTime' Name='ActualClosureDateCorrection' StaticName='ActualClosureDateCorrection' DisplayName='ActualClosureDateCorrection' Format='DateOnly' />`
    },
    {
        internalName: "Remarks",
        schemaXml: `<Field Type='Note' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' NumLines='6' RichText='FALSE' />`
    },
    {
        internalName: "TypeOfAction",
        schemaXml: buildChoiceFieldSchema("TypeOfAction", "TypeOfAction", typeOfActionChoices as readonly string[])
    },
    {
        internalName: "PlannedClosureDatePreventive",
        schemaXml: `<Field Type='DateTime' Name='PlannedClosureDatePreventive' StaticName='PlannedClosureDatePreventive' DisplayName='PlannedClosureDatePreventive' Format='DateOnly' />`
    },
    {
        internalName: "ActualClosureDatePreventive",
        schemaXml: `<Field Type='DateTime' Name='ActualClosureDatePreventive' StaticName='ActualClosureDatePreventive' DisplayName='ActualClosureDatePreventive' Format='DateOnly' />`
    },
    {
        internalName: "RelatedSubMetric",
        schemaXml: `<Field Type='Text' Name='RelatedSubMetric' StaticName='RelatedSubMetric' DisplayName='RelatedSubMetric' MaxLength='255' />`
    }
] as const;

const defaultViewFields: readonly RootCauseAnalysisViewField[] = [
    "ID",
    "ProblemStatementNumber",
    "CauseCategory",
    "Cause",
    "RCASource",
    "RCAPriority",
    "RootCause",
    "ResponsibilityCorrection",
    "ResponsibilityCorrective",
    "ResponsibilityPreventive",
    "PlannedClosureDateCorrection",
    "ActualClosureDateCorrective",
    "PlannedClosureDatePreventive",
    "ActualClosureDatePreventive",
    "RelatedSubMetric",
    "ActualClosureDateCorrection",
    "Remarks",
    "TypeOfAction",
    "Attachments",
    "Modified",
    "Editor"
] as const;

const definition: ListProvisionDefinition<RootCauseAnalysisFieldName, RootCauseAnalysisViewField> = {
    title: LIST_TITLE,
    description: "Root Cause Analysis list",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionRootCauseAnalysis(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
    await sp.web.lists.getByTitle(LIST_TITLE).fields.getByInternalNameOrTitle("Title").update({ Title: "Problem Statement" });
}

export default provisionRootCauseAnalysis;
