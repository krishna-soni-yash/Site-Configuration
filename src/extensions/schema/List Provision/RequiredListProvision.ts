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
    FacilitationReports: "FacilitationReports",
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    const { provisionLlBpRc } = await import('./lists/LlBpRc');
    //const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricLogs');
    //const { provisionEmailLogs } = await import('./lists/EmailLogs');
    //const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');
    //const { provisionManagementTaskLog } = await import('./lists/ManagementTaskLog');
    //const { provisionManagementEffortLog } = await import('./lists/ManagementEffortLog');
    //const { provisionFacilitationReport } = await import('./lists/FacilitationReport');

    await provisionLlBpRc(sp);
    //await provisionProjectMetricLogs(sp);
    //await provisionEmailLogs(sp);
    //await provisionProjectMetrics(sp);
    //await provisionManagementTaskLog(sp);
    //await provisionManagementEffortLog(sp);
    //await provisionFacilitationReport(sp);
}
