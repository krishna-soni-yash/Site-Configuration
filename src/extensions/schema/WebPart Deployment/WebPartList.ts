export interface IWebPartEntry {
	id: string;
	pageName: string;
	homePage?: boolean;
}

enum WebParts{
	HomePageWebPart = 'c2bf64bf-1b62-4aaf-9064-5b0793ce1829',
	MOMWebPart = '28449fdc-6c05-47a1-b58e-53d9911420a2',
	AuditWebPart = 'b244e89f-48d8-4cd7-b447-c9c5b4ccafa9',
	MetricsDashBoardWebPart = 'a9feb4c3-cca4-418d-b6b6-28a4263c8f79'
}

export const WebPartList: IWebPartEntry[] = [
	{ id: WebParts.HomePageWebPart, pageName: 'MOM', homePage: false },
	{ id: WebParts.MOMWebPart, pageName: 'MOM', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'Audits', homePage: false },
	{ id: WebParts.AuditWebPart, pageName: 'Audits', homePage: false },

	{ id: WebParts.HomePageWebPart, pageName: 'MetricsDashboard', homePage: false },
	{ id: WebParts.MetricsDashBoardWebPart, pageName: 'MetricsDashboard', homePage: false }
];

export default WebPartList;