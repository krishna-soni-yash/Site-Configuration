/*eslint-disable*/
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/items";
import { fetchListId } from "./GenericListProvision";
import {
    CurrentSchemaVersion,
    normalizeSchemaVersion,
    shouldRunUpdatedSchemaProvision,
    UpdatedListsForSchemaProvision
} from "./SchemaVersion";

export const RequiredListsProvision = {
    ListSchemaVersion: "ListSchemaVersion",
    ProjectMetricLogs: "ProjectMetricLogs",
    EmailLogs: "EmailLogs",
    ProjectMetrics: "ProjectMetrics",
    LlBpRc: "LlBpRc",

    //Audit & Facilitation Lists
    ManagementTaskLog: "ManagementTaskLog",
    ManagementEffortLog: "ManagementEffortLog",
    FacilitationReport: "FacilitationReport",

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
    SDLCParams: "SDLCParams",

    RootCauseAnalysis: "RootCauseAnalysis",
    CustomerSatisfactionIndex: "Customer Satisfaction Index",
    WorkLogManagement: "WorkLogManagement",
    TaskManagement: "TaskManagement",
    CodeReviewDefects: "Code Review Defects",
    TestingDefects: "Testing Defects",
    ReviewDefects: "Review Defects",
    MonthlyWorkdays: "MonthlyWorkdays",
    SprintMaster: "SprintMaster",

    //Graphs Lists
    ApplicableGraphs: "ApplicableGraphs",
    ResourceUtilization: "ResourceUtilization",
    CostOfQuality: "CostOfQuality",
    ScheduleVariation: "ScheduleVariation",
    OverallProductivity: "OverallProductivity",
    EffortDistribution: "EffortDistribution",
    RAED: "RAED",
    CRDD: "CRDD",
    EffortVariation: "EffortVariation",
    PostDeliveryDefects: "PostDeliveryDefects",
    CodingProductivity: "CodingProductivity",
    InternalDefects: "InternalDefects",
    DefectDensity: "DefectDensity",
    CodeReviewEffortDensity: "CodeReviewEffortDensity",
    CodeReviewReworkEffortDensity: "CodeReviewReworkEffortDensity",
    UnitTestingEffortDensity: "UnitTestingEffortDensity",
    TestExecutionEffortDensity: "TestExecutionEffortDensity",
    RiskSummary: "RiskSummary",
    FindingsSummary: "FindingsSummary",
    AgingFindings: "AgingFindings",
    OpenRootCause: "OpenRootCause",
    OpenIssues: "OpenIssues",
    OpenActionItems: "OpenActionItems",
    SpillOverIndex: "SpillOverIndex",
    Velocity: "Velocity",
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    const { provisionListSchemaVersion } = await import('./lists/ListSchemaVersion');
    const { provisionApplicableGraphs } = await import('./lists/ApplicableGraphs');
    const { provisionLlBpRc } = await import('./lists/LlBpRc');
    const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricLogs');
    const { provisionEmailLogs } = await import('./lists/EmailLogs');
    const { provisionManagementTaskLog } = await import('./lists/ManagementTaskLog');
    const { provisionMinutesOfMeeting } = await import('./lists/MinutesOfMeeting');
    const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');
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
    const { provisionAMSTicketEffortLog } = await import('./lists/AMSTicketEffortLog');
    const { provisionEmailErrorLogs } = await import('./lists/EmailErrorLogs');
    const { provisionQualityActivities } = await import('./lists/QualityActivities');
    const { provisionCodeReviewDefects } = await import('./lists/CodeReviewDefects');
    const { provisionTestingDefects } = await import('./lists/TestingDefects');
    const { provisionManagementEffortLog } = await import('./lists/ManagementEffortLog');
    const { provisionFacilitationReport } = await import('./lists/FacilitationReport');
    const { provisionSDLCParams } = await import('./lists/SDLCParams');
    const { provisionReviewDefects } = await import('./lists/ReviewDefects');
    const { provisionResourceUtilization } = await import('./lists/ResourceUtilization');
    const { provisionCostOfQuality } = await import('./lists/CostOfQuality');
    const { provisionScheduleVariation } = await import('./lists/ScheduleVariation');
    const { provisionOverallProductivity } = await import('./lists/OverallProductivity');
    const { provisionEffortDistribution } = await import('./lists/EffortDistribution');
    const { provisionRAED } = await import('./lists/RAED');
    const { provisionCRDD } = await import('./lists/CRDD');
    const { provisionEffortVariation } = await import('./lists/EffortVariation');
    const { provisionPostDeliveryDefects } = await import('./lists/PostDeliveryDefects');
    const { provisionCodingProductivity } = await import('./lists/CodingProductivity');
    const { provisionInternalDefects } = await import('./lists/InternalDefects');
    const { provisionDefectDensity } = await import('./lists/DefectDensity');
    const { provisionCodeReviewEffortDensity } = await import('./lists/CodeReviewEffortDensity');
    const { provisionCodeReviewReworkEffortDensity } = await import('./lists/CodeReviewReworkEffortDensity');
    const { provisionUnitTestingEffortDensity } = await import('./lists/UnitTestingEffortDensity');
    const { provisionTestExecutionEffortDensity } = await import('./lists/TestExecutionEffortDensity');
    const { provisionMonthlyWorkdays } = await import('./lists/MonthlyWorkdays');
    const { provisionRiskSummary } = await import('./lists/RiskSummary');
    const { provisionFindingsSummary } = await import('./lists/FindingsSummary');
    const { provisionAgingFindings } = await import('./lists/AgingFindings');
    const { provisionOpenRootCause } = await import('./lists/OpenRootCause');
    const { provisionOpenIssues } = await import('./lists/OpenIssues');
    const { provisionOpenActionItems } = await import('./lists/OpenActionItems');
    const { provisionSpillOverIndex } = await import('./lists/SpillOverIndex');
    const { provisionVelocity } = await import('./lists/Velocity');
    const { provisionSprintMaster } = await import('./lists/SprintMaster');

    await provisionListSchemaVersion(sp);

    const schemaVersionList = sp.web.lists.getByTitle(RequiredListsProvision.ListSchemaVersion);
    const schemaVersionItems = await schemaVersionList.items
        .select("ID", "Title")
        .orderBy("ID", false)
        .top(1)();

    const shouldRunAllSelectedLists = schemaVersionItems.length === 0;

    if (!shouldRunAllSelectedLists) {
        const lastSchemaVersionEntry = schemaVersionItems[0];
        const lastSchemaVersionFromTitle = normalizeSchemaVersion(lastSchemaVersionEntry?.Title);

        if (lastSchemaVersionFromTitle === CurrentSchemaVersion) {
            return;
        }

        const shouldRunUpdatedListsOnly = shouldRunUpdatedSchemaProvision(lastSchemaVersionFromTitle);

        if (!shouldRunUpdatedListsOnly) {
            return;
        }

        await provisionApplicableGraphs(sp);
        const applicableGraphsListIdForUpdates = await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);

        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.LlBpRc)) provisionLlBpRc(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ProjectMetricLogs)) provisionProjectMetricLogs(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.EmailLogs)) provisionEmailLogs(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ManagementTaskLog)) provisionManagementTaskLog(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.MinutesOfMeeting)) provisionMinutesOfMeeting(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ProjectMetrics)) provisionProjectMetrics(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ActionItemsTracker)) provisionActionItemsTracker(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.AdjustmentFactorValue)) provisionAdjustmentFactorValue(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ComplexityWeightage)) provisionComplexityWeightage(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.AMSMTTR)) provisionAMSMTTR(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ImpactValue)) provisionImpactValue(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.PotentialCost)) provisionPotentialCost(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ProbabilityValue)) provisionProbabilityValue(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.RAIDDescription)) provisionRAIDDescription(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.RAIDLogs)) provisionRAIDLogs(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.RootCauseAnalysis)) provisionRootCauseAnalysis(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CustomerSatisfactionIndex)) provisionCustomerSatisfactionIndex(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.WorkLogManagement)) provisionWorkLogManagement(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.PotentialBenefit)) provisionPotentialBenefit(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.TaskManagement)) provisionTaskManagement(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.AMSTicketLog)) provisionAMSTicketLog(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.EmailErrorLogs)) provisionEmailErrorLogs(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.AMSTicketEffortLog)) provisionAMSTicketEffortLog(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.QualityActivities)) provisionQualityActivities(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CodeReviewDefects)) provisionCodeReviewDefects(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.TestingDefects)) provisionTestingDefects(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ManagementEffortLog)) provisionManagementEffortLog(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.FacilitationReport)) provisionFacilitationReport(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.SDLCParams)) provisionSDLCParams(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ReviewDefects)) provisionReviewDefects(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ResourceUtilization)) provisionResourceUtilization(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CostOfQuality)) provisionCostOfQuality(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.ScheduleVariation)) provisionScheduleVariation(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.OverallProductivity)) provisionOverallProductivity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.EffortDistribution)) provisionEffortDistribution(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.RAED)) provisionRAED(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CRDD)) provisionCRDD(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.EffortVariation)) provisionEffortVariation(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.PostDeliveryDefects)) provisionPostDeliveryDefects(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CodingProductivity)) provisionCodingProductivity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.InternalDefects)) provisionInternalDefects(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.DefectDensity)) provisionDefectDensity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CodeReviewEffortDensity)) provisionCodeReviewEffortDensity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.CodeReviewReworkEffortDensity)) provisionCodeReviewReworkEffortDensity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.UnitTestingEffortDensity)) provisionUnitTestingEffortDensity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.TestExecutionEffortDensity)) provisionTestExecutionEffortDensity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.MonthlyWorkdays)) provisionMonthlyWorkdays(sp);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.RiskSummary)) provisionRiskSummary(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.FindingsSummary)) provisionFindingsSummary(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.AgingFindings)) provisionAgingFindings(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.OpenRootCause)) provisionOpenRootCause(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.OpenIssues)) provisionOpenIssues(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.OpenActionItems)) provisionOpenActionItems(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.SpillOverIndex)) provisionSpillOverIndex(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.Velocity)) provisionVelocity(sp, applicableGraphsListIdForUpdates);
        if (UpdatedListsForSchemaProvision.has(RequiredListsProvision.SprintMaster)) provisionSprintMaster(sp);

        schemaVersionList.items
            .getById(lastSchemaVersionEntry.ID)
            .update({ Title: `${CurrentSchemaVersion}` });

        return;
    }

    await provisionApplicableGraphs(sp);
    const applicableGraphsListId = await fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
    provisionLlBpRc(sp);
    provisionProjectMetricLogs(sp);
    provisionEmailLogs(sp);
    provisionManagementTaskLog(sp);
    provisionMinutesOfMeeting(sp);
    provisionProjectMetrics(sp);
    provisionActionItemsTracker(sp);
    provisionAdjustmentFactorValue(sp);
    provisionComplexityWeightage(sp);
    provisionAMSMTTR(sp);
    provisionImpactValue(sp);
    provisionPotentialCost(sp);
    provisionProbabilityValue(sp);
    provisionRAIDDescription(sp);
    provisionRAIDLogs(sp);
    provisionRootCauseAnalysis(sp);
    provisionCustomerSatisfactionIndex(sp);
    provisionWorkLogManagement(sp);
    provisionPotentialBenefit(sp);
    provisionTaskManagement(sp);
    provisionAMSTicketLog(sp);
    provisionEmailErrorLogs(sp);
    provisionAMSTicketEffortLog(sp);
    provisionQualityActivities(sp);
    provisionCodeReviewDefects(sp);
    provisionTestingDefects(sp);
    provisionManagementEffortLog(sp);
    provisionFacilitationReport(sp);
    provisionSDLCParams(sp);
    provisionReviewDefects(sp);
    provisionResourceUtilization(sp, applicableGraphsListId);
    provisionCostOfQuality(sp, applicableGraphsListId);
    provisionScheduleVariation(sp, applicableGraphsListId);
    provisionOverallProductivity(sp, applicableGraphsListId);
    provisionEffortDistribution(sp, applicableGraphsListId);
    provisionRAED(sp, applicableGraphsListId);
    provisionCRDD(sp, applicableGraphsListId);
    provisionEffortVariation(sp, applicableGraphsListId);
    provisionPostDeliveryDefects(sp, applicableGraphsListId);
    provisionCodingProductivity(sp, applicableGraphsListId);
    provisionInternalDefects(sp, applicableGraphsListId);
    provisionDefectDensity(sp, applicableGraphsListId);
    provisionCodeReviewEffortDensity(sp, applicableGraphsListId);
    provisionCodeReviewReworkEffortDensity(sp, applicableGraphsListId);
    provisionUnitTestingEffortDensity(sp, applicableGraphsListId);
    provisionTestExecutionEffortDensity(sp, applicableGraphsListId);
    provisionMonthlyWorkdays(sp);
    provisionRiskSummary(sp, applicableGraphsListId);
    provisionFindingsSummary(sp, applicableGraphsListId);
    provisionAgingFindings(sp, applicableGraphsListId);
    provisionOpenRootCause(sp, applicableGraphsListId);
    provisionOpenIssues(sp, applicableGraphsListId);
    provisionOpenActionItems(sp, applicableGraphsListId);
    provisionSpillOverIndex(sp, applicableGraphsListId);
    provisionVelocity(sp, applicableGraphsListId);
    provisionSprintMaster(sp);

    try {
        schemaVersionList.items.add({
            Title: `${CurrentSchemaVersion}`
        });
    } catch {
        const latestSchemaItem = schemaVersionList.items
            .select("ID")
            .orderBy("ID", false)
            .top(1)();

        if ((await latestSchemaItem).length > 0) {
            schemaVersionList.items
                .getById((await latestSchemaItem)[0].ID)
                .update({ Title: `${CurrentSchemaVersion}` });
        }
    }
}
