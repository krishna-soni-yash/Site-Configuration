/*eslint-disable*/
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

export const RequiredListsProvision = {
    ProjectMetricLogs: "ProjectMetricLogs",
    EmailLogs: "EmailLogs",
    ProjectMetrics: "ProjectMetrics",
    LlBpRc: "LlBpRc",

    //Audit & Facilitation Lists
    ManagementTaskLog: "ManagementTaskLog",
    ManagementEffortLog: "ManagementEffortLog",
    FacilitationReport: "FacilitationReports",

    //MOM & Action Items List
    MinutesOfMeeting: "MinutesOfMeeting",
    ActionItemsTracker: "ActionItemsTracker",

    AdjustmentFactorValue: "AdjustmentFactorValue",
    AMSMTTR: "AMSMTTR",
    ComplexityWeightage: "ComplexityWeightage",
    ImpactValue: "ImpactValue",
    PotentialBenefit: "PotentialBenefit",
    PotentialCost: "PotentialCost",
    RAIDLogs: "RAIDLogs",
    ProbabilityValue: "ProbabilityValue",
    RAIDDescription: "RAIDDescription",
    AMSTicketLog: "AMSTicketLog",
    AMSTicketEffortLog: "AMSTicketEffortLog",
    EmailErrorLogs: "EmailErrorLogs",
    QualityActivities: "QualityActivities",

    RootCauseAnalysis: "RootCauseAnalysis",
    CustomerSatisfactionIndex: "Customer Satisfaction Index",
    WorkLogManagement: "WorkLogManagement",
    TaskManagement: "TaskManagement",
    CodeReviewDefects: "CodeReviewDefects",
    TestingDefects: "Testing Defects",
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    const { provisionLlBpRc } = await import('./lists/LlBpRc');
    const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricLogs');
    const { provisionEmailLogs } = await import('./lists/EmailLogs');
    const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');
    const { provisionManagementTaskLog } = await import('./lists/ManagementTaskLog');
    const { provisionManagementEffortLog } = await import('./lists/ManagementEffortLog');
    const { provisionFacilitationReport } = await import('./lists/FacilitationReport');
    const { provisionMinutesOfMeeting } = await import('./lists/MinutesOfMeeting');
    const { provisionActionItemsTracker } = await import('./lists/ActionItemsTracker');
    const { provisionAdjustmentFactorValue } = await import('./lists/AdjustmentFactorValue');
    const { provisionAMSMTTR } = await import('./lists/AMSMTTR');
    const { provisionComplexityWeightage } = await import('./lists/ComplexityWeightage');
    const { provisionImpactValue } = await import('./lists/ImpactValue');
    const { provisionPotentialBenefit } = await import('./lists/PotentialBenefit');
    const { provisionPotentialCost } = await import('./lists/PotentialCost');
    const { provisionProbabilityValue } = await import('./lists/ProbabilityValue');
    const { provisionRAIDDescription } = await import('./lists/RAIDDescription');
    const { provisionRAIDLogs } = await import('./lists/RAIDLogs');
    const { provisionRootCauseAnalysis } = await import('./lists/RootCauseAnalysis');
    const { provisionCustomerSatisfactionIndex } = await import('./lists/CustomerSatisfactionIndex');
    const { provisionWorkLogManagement } = await import('./lists/WorkLogManagement');
    const { provisionTaskManagement } = await import('./lists/TaskManagement');
    const { provisionAMSTicketLog } = await import('./lists/AMSTicketLog');
    const {provisionAMSTicketEffortLog} = await import('./lists/AMSTicketEffortLog');
    const { provisionEmailErrorLogs } = await import('./lists/EmailErrorLogs');
    const { provisionQualityActivities } = await import('./lists/QualityActivities');
    const { provisionCodeReviewDefects } = await import('./lists/CodeReviewDefects');
    const { provisionTestingDefects } = await import('./lists/TestingDefects');

    provisionLlBpRc(sp);
    provisionProjectMetricLogs(sp);
    provisionEmailLogs(sp);
    provisionProjectMetrics(sp);
    provisionManagementTaskLog(sp);
    provisionManagementEffortLog(sp);
    provisionFacilitationReport(sp);
    provisionMinutesOfMeeting(sp);
    provisionActionItemsTracker(sp);
    provisionAdjustmentFactorValue(sp);
    provisionAMSMTTR(sp);
    provisionComplexityWeightage(sp);
    provisionImpactValue(sp);
    provisionPotentialBenefit(sp);
    provisionPotentialCost(sp);
    provisionProbabilityValue(sp);
    provisionRAIDDescription(sp);
    provisionRAIDLogs(sp);
    provisionRootCauseAnalysis(sp);
    provisionCustomerSatisfactionIndex(sp);
    provisionWorkLogManagement(sp);
    provisionTaskManagement(sp);
    provisionAMSTicketLog(sp);
    provisionAMSTicketEffortLog(sp);
    provisionEmailErrorLogs(sp);
    provisionQualityActivities(sp);
    provisionCodeReviewDefects(sp);
    provisionCodeReviewDefects(sp);
    provisionTestingDefects(sp);
}
