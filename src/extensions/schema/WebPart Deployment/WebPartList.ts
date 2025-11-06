export interface IWebPartEntry {
	id: string;
	pageName: string;
	homePage?: boolean;
}

export const WebPartList: IWebPartEntry[] = [
	{ id: 'aec7bd2e-5d17-4a98-89c0-ddb541197235', pageName: 'PPOQuality', homePage: true },
];

export default WebPartList;