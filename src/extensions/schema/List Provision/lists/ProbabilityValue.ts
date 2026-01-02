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

const LIST_TITLE = RequiredListsProvision.ProbabilityValue;

type ProbabilityValueFieldName = "Text";

type ProbabilityValueViewField = ProbabilityValueFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<ProbabilityValueFieldName>[] = [
	{
		internalName: "Text",
		schemaXml: `<Field Type='Text' Name='Text' StaticName='Text' DisplayName='Text' MaxLength='1024' />`
	}
] as const;

const defaultViewFields: readonly ProbabilityValueViewField[] = [
	"LinkTitle",
	"Text"
] as const;

const definition: ListProvisionDefinition<ProbabilityValueFieldName, ProbabilityValueViewField> = {
	title: LIST_TITLE,
	description: "Probability value lookup list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

const seedItems: ReadonlyArray<{ Title: string; Text: string }> = [
	{ Title: "1", Text: "Not likely to occur" },
	{ Title: "2", Text: "Not very likely to occur" },
	{ Title: "3", Text: "Somewhat less than an even chance" },
	{ Title: "4", Text: "An even chance to occur" },
	{ Title: "5", Text: "Somewhat greater than an even chance" },
	{ Title: "6", Text: "Likely to occur" },
	{ Title: "7", Text: "Very likely to occur" },
	{ Title: "8", Text: "Almost sure to occur" },
	{ Title: "9", Text: "Extremely sure to occur" },
	{ Title: "10", Text: "Certain to occur" }
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

export async function provisionProbabilityValue(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureSeedData(sp);
}

export default provisionProbabilityValue;
