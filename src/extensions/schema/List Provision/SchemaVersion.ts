export const CurrentSchemaVersion = 1;

export const SelectedListsForSchemaProvision: readonly string[] = [
	//"ProjectMetricLogs",
	//"EmailLogs",
	// "ProjectMetrics",
	// "LlBpRc",
	// "ManagementTaskLog",
	// "ManagementEffortLog",
	// "FacilitationReport",
	// "MinutesOfMeeting",
	// "ActionItemsTracker",
	// "AdjustmentFactorValue",
	// "AMSMTTR",
	// "ComplexityWeightage",
	// "ImpactValue",
	// "PotentialBenefit",
	// "PotentialCost",
	// "RAIDLogs",
	// "ProbabilityValue",
	// "RAIDDescription",
	// "AMSTicketLog",
	// "AMSTicketEffortLog",
	// "EmailErrorLogs",
	// "QualityActivities",
	// "SDLCParams",
	// "RootCauseAnalysis",
	// "Customer Satisfaction Index",
	// "WorkLogManagement",
	// "TaskManagement",
	// "Code Review Defects",
	// "Testing Defects",
	// "Review Defects",
	// "MonthlyWorkdays",
	// "ApplicableGraphs",
	// "ResourceUtilization",
	// "CostOfQuality",
	// "ScheduleVariation",
	// "OverallProductivity",
	// "EffortDistribution",
	// "RAED",
	// "CRDD",
	// "EffortVariation",
	// "PostDeliveryDefects",
	// "CodingProductivity",
	// "InternalDefects",
	// "DefectDensity",
	// "CodeReviewEffortDensity",
	// "CodeReviewReworkEffortDensity",
	// "UnitTestingEffortDensity",
	// "TestExecutionEffortDensity",
	// "RiskSummary",
	// "FindingsSummary",
	// "AgingFindings",
	// "OpenRootCause",
	// "OpenIssues",
	// "OpenActionItems",
	//"SpillOverIndex",
	//"Velocity",
	//"SprintMaster"
] as const;

export const UpdatedListsForSchemaProvision: ReadonlySet<string> = new Set<string>([
	...SelectedListsForSchemaProvision
]);

export function normalizeSchemaVersion(value: unknown): number {
	const parsed = Number(`${value ?? ""}`.trim());
	return Number.isFinite(parsed) ? parsed : 0;
}

export function shouldRunUpdatedSchemaProvision(lastRecordedSchemaVersion: number): boolean {
	return CurrentSchemaVersion > lastRecordedSchemaVersion;
}

