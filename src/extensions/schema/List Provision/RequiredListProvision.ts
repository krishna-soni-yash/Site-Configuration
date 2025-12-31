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

    RootCauseAnalysis: "RootCauseAnalysis",
    CustomerSatisfactionIndex: "Customer Satisfaction Index",
    WorkLogManagement: "WorkLogManagement",
    TaskManagement: "TaskManagement",
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    //const { provisionLlBpRc } = await import('./lists/LlBpRc');
    //const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricLogs');
    //const { provisionEmailLogs } = await import('./lists/EmailLogs');
    //const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');
    //const { provisionManagementTaskLog } = await import('./lists/ManagementTaskLog');
    //const { provisionManagementEffortLog } = await import('./lists/ManagementEffortLog');
    //const { provisionFacilitationReport } = await import('./lists/FacilitationReport');
    //const { provisionMinutesOfMeeting } = await import('./lists/MinutesOfMeeting');
    //const { provisionActionItemsTracker } = await import('./lists/ActionItemsTracker');
    //const { provisionAdjustmentFactorValue } = await import('./lists/AdjustmentFactorValue');
    //const { provisionAMSMTTR } = await import('./lists/AMSMTTR');
    //const { provisionComplexityWeightage } = await import('./lists/ComplexityWeightage');
    //const { provisionImpactValue } = await import('./lists/ImpactValue');
    //const { provisionPotentialBenefit } = await import('./lists/PotentialBenefit');
    //const { provisionPotentialCost } = await import('./lists/PotentialCost');
    //const { provisionProbabilityValue } = await import('./lists/ProbabilityValue');
    //const { provisionRAIDDescription } = await import('./lists/RAIDDescription');
    //const { provisionRAIDLogs } = await import('./lists/RAIDLogs');
    //const { provisionRootCauseAnalysis } = await import('./lists/RootCauseAnalysis');
    //const { provisionCustomerSatisfactionIndex } = await import('./lists/CustomerSatisfactionIndex');
    //const { provisionWorkLogManagement } = await import('./lists/WorkLogManagement');
    //const { provisionTaskManagement } = await import('./lists/TaskManagement');
    const { provisionAMSTicketLog } = await import('./lists/AMSTicketLog');

    //await provisionLlBpRc(sp);
    //await provisionProjectMetricLogs(sp);
    //await provisionEmailLogs(sp);
    //await provisionProjectMetrics(sp);
    //await provisionManagementTaskLog(sp);
    //await provisionManagementEffortLog(sp);
    //await provisionFacilitationReport(sp);
    //await provisionMinutesOfMeeting(sp);
    //await provisionActionItemsTracker(sp);
    //await provisionAdjustmentFactorValue(sp);
    //await provisionAMSMTTR(sp);
    //await provisionComplexityWeightage(sp);
    //await provisionImpactValue(sp);
    //await provisionPotentialBenefit(sp);
    //await provisionPotentialCost(sp);
    //await provisionProbabilityValue(sp);
    //await provisionRAIDDescription(sp);
    //await provisionRAIDLogs(sp);
    //await provisionRootCauseAnalysis(sp);
    //await provisionCustomerSatisfactionIndex(sp);
    //await provisionWorkLogManagement(sp);
    //await provisionTaskManagement(sp);
    await provisionAMSTicketLog(sp);
}
