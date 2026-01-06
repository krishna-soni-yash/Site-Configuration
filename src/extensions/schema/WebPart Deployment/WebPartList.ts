export interface IWebPartEntry {
	id: string;
	pageName: string;
	homePage?: boolean;
}

// enum WebParts{
// 	HomePageWebPart = 'c2bf64bf-1b62-4aaf-9064-5b0793ce1829',
// 	MOMWebPart = '28449fdc-6c05-47a1-b58e-53d9911420a2',
// 	AuditWebPart = 'b244e89f-48d8-4cd7-b447-c9c5b4ccafa9',
// 	MetricsDashBoardWebPart = 'a9feb4c3-cca4-418d-b6b6-28a4263c8f79',
// 	RootCauseAnalysisWebPart = 'b77d3069-e9d7-4521-a93c-a7ec0b2dfa50',
// 	PPOWebPart = 'aec7bd2e-5d17-4a98-89c0-ddb541197235',
// 	AMSWebpart = 'bc0d2a5f-1168-4a7b-b218-6b39bffd9e11',
// 	EstimationsWebPart = '1e0bf94d-fcd7-4556-97ac-a334deb8ff36'
// }

export const WebPartList: IWebPartEntry[] = [
	// { id: WebParts.HomePageWebPart, pageName: 'Home', homePage: true },
	// { id: WebParts.MetricsDashBoardWebPart, pageName: 'Home', homePage: true },

	// { id: WebParts.HomePageWebPart, pageName: 'MoM-ActionItem', homePage: false },
	// { id: WebParts.MOMWebPart, pageName: 'MoM-ActionItem', homePage: false },

	// { id: WebParts.HomePageWebPart, pageName: 'Audit', homePage: false },
	// { id: WebParts.AuditWebPart, pageName: 'Audit', homePage: false },

	// { id: WebParts.HomePageWebPart, pageName: 'RCA-And-Raid-Logs', homePage: false },
	// { id: WebParts.RootCauseAnalysisWebPart, pageName: 'RCA-And-Raid-Logs', homePage: false },

	// { id: WebParts.HomePageWebPart, pageName: 'PPO', homePage: false },
	// { id: WebParts.PPOWebPart, pageName: 'PPO', homePage: false },

	// { id: WebParts.HomePageWebPart, pageName: 'AMS', homePage: false },
	// { id: WebParts.AMSWebpart, pageName: 'AMS', homePage: false },

	// { id: WebParts.HomePageWebPart, pageName: 'Estimations', homePage: false },
	// { id: WebParts.EstimationsWebPart, pageName: 'Estimations', homePage: false }
];

export default WebPartList;