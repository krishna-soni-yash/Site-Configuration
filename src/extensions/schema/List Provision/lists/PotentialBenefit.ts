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

const LIST_TITLE = RequiredListsProvision.PotentialBenefit;

type PotentialBenefitFieldName = "Text";

type PotentialBenefitViewField = PotentialBenefitFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<PotentialBenefitFieldName>[] = [
	{
		internalName: "Text",
		schemaXml: `<Field Type='Text' Name='Text' StaticName='Text' DisplayName='Text' MaxLength='1024' />`
	}
] as const;

const defaultViewFields: readonly PotentialBenefitViewField[] = [
	"LinkTitle",
	"Text"
] as const;

const definition: ListProvisionDefinition<PotentialBenefitFieldName, PotentialBenefitViewField> = {
	title: LIST_TITLE,
	description: "Potential benefit lookup list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

const seedItems: ReadonlyArray<{ Title: string; Text: string }> = [
	{ Title: "1", Text: "No Benefit" },
	{ Title: "2", Text: "Very Low Benefit" },
	{ Title: "3", Text: "Low Benefit" },
	{ Title: "4", Text: "Below Moderate Benefit" },
	{ Title: "5", Text: "Moderate Benefit" },
	{ Title: "6", Text: "Above Moderate Benefit" },
	{ Title: "7", Text: "High Benefit" },
	{ Title: "8", Text: "Above High Benefit" },
	{ Title: "9", Text: "Very High Benefit" },
	{ Title: "10", Text: "Extreme High Benefit" }
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

export async function provisionPotentialBenefit(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureSeedData(sp);
}

export default provisionPotentialBenefit;
