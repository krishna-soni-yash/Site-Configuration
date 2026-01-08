/* eslint-disable */
import { SPFI } from "@pnp/sp";

import {
	DocumentLibraryProvisionDefinition,
	ensureDocumentLibrary
} from "../genericLibraryProvision";

const PROJECT_DOCUMENTS_FOLDERS = [
	"Customer",
	"Project Management",
	"Software Engineering"
] as const;

const PROJECT_MANAGEMENT_SUBFOLDERS = [
	"Project Management/01 Project Initiation",
	"Project Management/02 Project Planning",
	"Project Management/03 Project Monitoring and Control",
	"Project Management/04 Project Closure",
    "Software Engineering/01 Requirenments",
    "Software Engineering/02 Design",
    "Software Engineering/03 Development",
    "Software Engineering/04 Testing",
    "Software Engineering/05 Delivery",
] as const;

type ProjectDocumentsFieldName = string;
type ProjectDocumentsViewField = ProjectDocumentsFieldName;

const projectDocumentsDefinition: DocumentLibraryProvisionDefinition<
	ProjectDocumentsFieldName,
	ProjectDocumentsViewField
> = {
	title: "Project Documents",
	description: "Stores project related documentation and assets.",
	enableVersioning: true,
	enableFolderCreation: true,
	folders: {
		folderPaths: [...PROJECT_DOCUMENTS_FOLDERS, ...PROJECT_MANAGEMENT_SUBFOLDERS]
	}
};

export async function provisionProjectDocumentsLibrary(sp: SPFI): Promise<void> {
	await ensureDocumentLibrary(sp, projectDocumentsDefinition);
}

export const ProjectDocumentsLibraryName = projectDocumentsDefinition.title;
