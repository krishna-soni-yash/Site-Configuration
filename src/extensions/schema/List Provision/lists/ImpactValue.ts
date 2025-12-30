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

const LIST_TITLE = RequiredListsProvision.ImpactValue;

type ImpactValueFieldName = "Text";

type ImpactValueViewField = ImpactValueFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<ImpactValueFieldName>[] = [
	{
		internalName: "Text",
		schemaXml: `<Field Type='Text' Name='Text' StaticName='Text' DisplayName='Text' MaxLength='1024' />`
	}
] as const;

const defaultViewFields: readonly ImpactValueViewField[] = [
	"LinkTitle",
	"Text"
] as const;

const definition: ListProvisionDefinition<ImpactValueFieldName, ImpactValueViewField> = {
	title: LIST_TITLE,
	description: "Impact value lookup list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

const seedItems: ReadonlyArray<{ Title: string; Text: string }> = [
	{ Title: "1", Text: "Very Low Impact" },
	{ Title: "2", Text: "Low Impact - Needs No Attention" },
	{
		Title: "3",
		Text: "Has a minor impact and things are taken into considerations by the respective projects / departments"
	},
	{ Title: "4", Text: "Medium impact and needs attention for resolving the issue" },
	{ Title: "5", Text: "Has an impact and resolves in co-ordination with BUH / CEO- and DH" },
	{ Title: "6", Text: "Has an high impact and which effects business value" },
	{ Title: "7", Text: "Has an major impact and which has significant effect on business" },
	{ Title: "8", Text: "Has a critical impact and call for a review with senior management" },
	{ Title: "9", Text: "Has a very critical impact and needs senior management attention / CISO" },
	{ Title: "10", Text: "Leads to serious problems and call for emergency mode" }
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

export async function provisionImpactValue(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureSeedData(sp);
}

export default provisionImpactValue;
