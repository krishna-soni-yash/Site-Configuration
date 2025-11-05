import { spfi, SPFx, SPFI } from "@pnp/sp";

export const RequiredListsProvision = {
    ProjectMetricLogs: "ProjectMetricLogs",
    EmailLogs: "EmailLogs",
    ProjectMetrics: "ProjectMetrics"
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricLogs');
    const { provisionEmailLogs } = await import('./lists/EmailLogs');
    const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');

    await provisionProjectMetricLogs(sp);
    await provisionEmailLogs(sp);
    await provisionProjectMetrics(sp);
}
