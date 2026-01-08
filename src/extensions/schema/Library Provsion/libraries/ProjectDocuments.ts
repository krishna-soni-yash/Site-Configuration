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
		folderPaths: [...PROJECT_DOCUMENTS_FOLDERS]
	}
};

export async function provisionProjectDocumentsLibrary(sp: SPFI): Promise<void> {
	await ensureDocumentLibrary(sp, projectDocumentsDefinition);
}

export const ProjectDocumentsLibraryName = projectDocumentsDefinition.title;
