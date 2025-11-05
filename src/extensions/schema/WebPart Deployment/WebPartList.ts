export interface IWebPartEntry {
	id: string;
	pageName: string;
	homePage?: boolean;
}

export const WebPartList: IWebPartEntry[] = [
	//{ id: 'aec7bd2e-5d17-4a98-89c0-ddb541197235', pageName: 'PPOQuality', homePage: true },
	{ id: '59051b24-75c9-44bc-83b8-7fe01824981a', pageName: 'RequiredList', homePage: false },
];

export default WebPartList;