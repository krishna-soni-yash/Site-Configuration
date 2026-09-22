/*eslint-disable*/
import { ProjectDocumentsLibraryName } from '../Library Provsion/libraries/ProjectDocuments';

export interface IListBindingConfig {
	listTitle: string;
	listId?: string;
	viewTitle?: string;
	viewId?: string;
	webPartTitle?: string;
	isDocumentLibrary?: boolean;
}

export interface IWebPartEntry {
	id?: string;
	alias?: string;
	pageName: string;
	homePage?: boolean;
	listBinding?: IListBindingConfig;
}

enum WebParts{
 	HomePageWebPart = 'c2bf64bf-1b62-4aaf-9064-5b0793ce1829',
	MOMWebPart = '28449fdc-6c05-47a1-b58e-53d9911420a2',
	AuditWebPart = 'b244e89f-48d8-4cd7-b447-c9c5b4ccafa9',
 	MetricsDashBoardWebPart = '75e7d5ba-b4ae-42cd-a3ba-cff52f5f4056',
	RootCauseAnalysisWebPart = 'b77d3069-e9d7-4521-a93c-a7ec0b2dfa50',
	PPOWebPart = 'aec7bd2e-5d17-4a98-89c0-ddb541197235',
	AMSWebpart = 'bc0d2a5f-1168-4a7b-b218-6b39bffd9e11',
	EstimationsWebPart = '1e0bf94d-fcd7-4556-97ac-a334deb8ff36',
	HideBannerWebPart = '137d8eae-8e43-42fd-ad03-52842b02daeb',
	DefectsWebPart = 'd199e076-7cf1-4458-a274-eef36888c787',
	CustomerSatisfactionIndexWebPart = '160bd2bf-b186-466d-9710-9a57a652cdca',
	FooterWebPart = '9b6f63a3-953a-4b11-96b0-b05763fb3179'
}

export const WebPartList: IWebPartEntry[] = [
	{ id: WebParts.HomePageWebPart, pageName: 'Audit', homePage: false },
	{ id: WebParts.AuditWebPart, pageName: 'Audit', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'Audit', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'Defects-Lists', homePage: false },
	{ id: WebParts.DefectsWebPart, pageName: 'Defects-Lists', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'Defects-Lists', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'Documents', homePage: false },
	{ id: WebParts.HideBannerWebPart, pageName: 'Documents', homePage: false },
	{
		alias: 'ListWebPart',
		pageName: 'Documents',
		listBinding: {
			listTitle: ProjectDocumentsLibraryName,
			webPartTitle: 'Project Documents',
			isDocumentLibrary: true
		},
	},
	{ id: WebParts.FooterWebPart, pageName: 'Documents', homePage: false },
	
	{ id: WebParts.HomePageWebPart, pageName: 'Home', homePage: true },
	//{ id: WebParts.HideBannerWebPart, pageName: 'Home', homePage: true },
	{ id: WebParts.MetricsDashBoardWebPart, pageName: 'Home', homePage: true },
	{ id: WebParts.FooterWebPart, pageName: 'Home', homePage: true },

	{ id: WebParts.HomePageWebPart, pageName: 'MoM-ActionItem', homePage: false },
	{ id: WebParts.MOMWebPart, pageName: 'MoM-ActionItem', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'MoM-ActionItem', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'RCA-And-Raid-Logs', homePage: false },
	{ id: WebParts.RootCauseAnalysisWebPart, pageName: 'RCA-And-Raid-Logs', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'RCA-And-Raid-Logs', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'PPO', homePage: false },
	{ id: WebParts.PPOWebPart, pageName: 'PPO', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'PPO', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'AMS', homePage: false },
	{ id: WebParts.AMSWebpart, pageName: 'AMS', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'AMS', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'Estimations', homePage: false },
	{ id: WebParts.EstimationsWebPart, pageName: 'Estimations', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'Estimations', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'CSI', homePage: false },
	{ id: WebParts.CustomerSatisfactionIndexWebPart, pageName: 'CSI', homePage: false },
	{ id: WebParts.FooterWebPart, pageName: 'CSI', homePage: false }
];

export default WebPartList;