import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/items";
import {
	ensureListProvision,
	FieldDefinition,
	ListProvisionDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.RAIDDescription;

type RAIDDescriptionFieldName = "Description";

type RAIDDescriptionViewField = RAIDDescriptionFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<RAIDDescriptionFieldName>[] = [
	{
		internalName: "Description",
		schemaXml: `<Field Type='Note' Name='Description' StaticName='Description' DisplayName='Description' NumLines='6' RichText='FALSE' />`
	}
] as const;

const defaultViewFields: readonly RAIDDescriptionViewField[] = [
	"LinkTitle",
	"Description"
] as const;

const definition: ListProvisionDefinition<RAIDDescriptionFieldName, RAIDDescriptionViewField> = {
	title: LIST_TITLE,
	description: "RAID description lookup list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

const seedItems: ReadonlyArray<{ Title: string; Description: string }> = [
	{ Title: "Risk", Description: "Lack of clarity in user specifications/ requirements." },
	{ Title: "Risk", Description: "The Azure Virtual Machine's performance falling below expected levels may negatively affect application performance and overall system efficiency." },
	{ Title: "Risk", Description: "Lack of proper understanding of the complexity of requirements" },
	{ Title: "Risk", Description: "Development of projects for clients where languages other than English are used" },
	{ Title: "Risk", Description: "Incomplete definition of scope / requirements" },
	{ Title: "Risk", Description: "Client representative is not available" },
	{ Title: "Risk", Description: "Excess project team members than required for the project" },
	{ Title: "Risk", Description: "Shortage of project team members with sufficient relevant experience" },
	{ Title: "Risk", Description: "Lack of project team members who have right combination of skills" },
	{ Title: "Risk", Description: "High attrition rates" },
	{ Title: "Risk", Description: "Lack of proper plan and schedule" },
	{ Title: "Risk", Description: "Lack of project team members commitment for the entire duration of project" },
	{ Title: "Risk", Description: "Multiple vendors / contractors" },
	{ Title: "Risk", Description: "Critical dependence on external suppliers" },
	{ Title: "Risk", Description: "Number of inter-project dependencies" },
	{ Title: "Risk", Description: "Plan requires extensive recruitment of personnel for project" },
	{ Title: "Risk", Description: "Multiple number of decision makers" },
	{ Title: "Risk", Description: "Multiple geographical locations" },
	{ Title: "Risk", Description: "Key users unavailable" },
	{ Title: "Risk", Description: "Dependent on scarce resources / limited skills" },
	{ Title: "Risk", Description: "Complex task dependencies" },
	{ Title: "Risk", Description: "Project Manager inexperience in handling similar projects" },
	{ Title: "Risk", Description: "Buffer time not provided to compensate delays due to unforeseen conditions" },
	{ Title: "Risk", Description: "Inaccurate estimation regarding size and complexity of the project" },
	{ Title: "Risk", Description: "Lack of proper design" },
	{ Title: "Risk", Description: "Technical/ technology  uncertainty in some or entire project development" },
	{ Title: "Risk", Description: "Inappropriate development tools" },
	{ Title: "Risk", Description: "New / unfamiliar technology" },
	{ Title: "Risk", Description: "Unstable development team (Frequent changes to project team members)" },
	{ Title: "Risk", Description: "Low team knowledge of business area" },
	{ Title: "Risk", Description: "Use of development method / standards" },
	{ Title: "Risk", Description: "Complex functions" },
	{ Title: "Risk", Description: "Complex database" },
	{ Title: "Risk", Description: "Database to be shared by a number of applications" },
	{ Title: "Risk", Description: "Number of physical system interfaces" },
	{ Title: "Risk", Description: "Numbers of design decisions at discretion of systems architect (no user involvement)" },
	{ Title: "Risk", Description: "Rapid response time (below 2 seconds)" },
	{ Title: "Risk", Description: "Small batch window" }
];

async function ensureSeedData(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const existing = await list.items.select("Id", "Title", "Description")();
	const existingKeys = new Map<string, number>();
	for (const item of existing) {
		const key = `${item.Title ?? ""}||${item.Description ?? ""}`;
		existingKeys.set(key, Number(item.Id));
	}

	const desiredKeys = new Set<string>();
	for (const seed of seedItems) {
		const key = `${seed.Title}||${seed.Description}`;
		desiredKeys.add(key);
		if (!existingKeys.has(key)) {
			await list.items.add({ Title: seed.Title, Description: seed.Description });
		}
	}

	const existingEntries = Array.from(existingKeys.entries());
	for (const [key, id] of existingEntries) {
		if (!desiredKeys.has(key)) {
			await list.items.getById(id).delete();
		}
	}
}

export async function provisionRAIDDescription(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureSeedData(sp);
}

export default provisionRAIDDescription;
