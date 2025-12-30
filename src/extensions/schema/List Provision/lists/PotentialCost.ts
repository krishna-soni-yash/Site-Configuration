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

const LIST_TITLE = RequiredListsProvision.PotentialCost;

type PotentialCostFieldName = "Text";

type PotentialCostViewField = PotentialCostFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<PotentialCostFieldName>[] = [
	{
		internalName: "Text",
		schemaXml: `<Field Type='Text' Name='Text' StaticName='Text' DisplayName='Text' MaxLength='1024' />`
	}
] as const;

const defaultViewFields: readonly PotentialCostViewField[] = [
	"LinkTitle",
	"Text"
] as const;

const definition: ListProvisionDefinition<PotentialCostFieldName, PotentialCostViewField> = {
	title: LIST_TITLE,
	description: "Potential cost lookup list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

const seedItems: ReadonlyArray<{ Title: string; Text: string }> = [
	{ Title: "1", Text: "No Cost" },
	{ Title: "2", Text: "Very Low Cost" },
	{ Title: "3", Text: "Low Cost" },
	{ Title: "4", Text: "Below Medium Cost" },
	{ Title: "5", Text: "Medium Cost" },
	{ Title: "6", Text: "Above Medium Cost" },
	{ Title: "7", Text: "High Cost" },
	{ Title: "8", Text: "Above High Cost" },
	{ Title: "9", Text: "Very High Cost" },
	{ Title: "10", Text: "Extreme High Cost" }
];

async function ensureSeedData(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const existing = await list.items.select("Id", "Title", "Text")();
	const existingMap = new Map<string, { id: number; text: string }>();
	for (const item of existing) {
		if (typeof item.Title === "string") {
			existingMap.set(item.Title, { id: Number(item.Id), text: `${item.Text ?? ""}` });
		}
	}

	for (const seed of seedItems) {
		const match = existingMap.get(seed.Title);
		if (!match) {
			await list.items.add({ Title: seed.Title, Text: seed.Text });
			continue;
		}
		if (match.text !== seed.Text) {
			await list.items.getById(match.id).update({ Text: seed.Text });
		}
	}
}

export async function provisionPotentialCost(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureSeedData(sp);
}

export default provisionPotentialCost;
