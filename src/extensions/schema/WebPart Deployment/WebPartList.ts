export interface IWebPartEntry {
	webPartIds: string[];
	pageName: string;
	homePage?: boolean;
}

export const WebPartList: IWebPartEntry[] = [
	{ webPartIds: ['b244e89f-48d8-4cd7-b447-c9c5b4ccafa9', 'c2bf64bf-1b62-4aaf-9064-5b0793ce1829'], pageName: 'AuditFacilitation', homePage: false },
];

export default WebPartList;