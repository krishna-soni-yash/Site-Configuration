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

const LIST_TITLE = RequiredListsProvision.ComplexityWeightage;

type ComplexityWeightageFieldName = "Simple" | "Medium" | "Complex" | "VeryComplex";

type ComplexityWeightageViewField = ComplexityWeightageFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<ComplexityWeightageFieldName>[] = [
	{
		internalName: "Simple",
		schemaXml: `<Field Type='Text' Name='Simple' StaticName='Simple' DisplayName='Simple' MaxLength='255' />`
	},
	{
		internalName: "Medium",
		schemaXml: `<Field Type='Text' Name='Medium' StaticName='Medium' DisplayName='Medium' MaxLength='255' />`
	},
	{
		internalName: "Complex",
		schemaXml: `<Field Type='Text' Name='Complex' StaticName='Complex' DisplayName='Complex' MaxLength='255' />`
	},
	{
		internalName: "VeryComplex",
		schemaXml: `<Field Type='Text' Name='VeryComplex' StaticName='VeryComplex' DisplayName='VeryComplex' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly ComplexityWeightageViewField[] = [
	"LinkTitle",
	"Simple",
	"Medium",
	"Complex",
	"VeryComplex"
] as const;

const definition: ListProvisionDefinition<ComplexityWeightageFieldName, ComplexityWeightageViewField> = {
	title: LIST_TITLE,
	description: "Complexity weightage list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

type SeedItem = {
	LinkTitle: string;
	Simple: number;
	Medium: number;
	Complex: number;
	VeryComplex: number;
};

const seedItems: readonly SeedItem[] = [
	{
		LinkTitle: "Agile",
		Simple: 1.1,
		Medium: 1.4,
		Complex: 1.6,
		VeryComplex: 3.2
	},
	{
		LinkTitle: "DEV",
		Simple: 0.9,
		Medium: 1.4,
		Complex: 2.6,
		VeryComplex: 4.7
	},
	{
		LinkTitle: "DEVM",
		Simple: 1,
		Medium: 1.2,
		Complex: 1.9,
		VeryComplex: 4.7
	}
];

async function ensureTitleRenamed(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	try {
		const field = await list.fields.getByInternalNameOrTitle("Title").select("Title")();
		if (!field || `${field.Title}` !== "ProjectType") {
			await list.fields.getByInternalNameOrTitle("Title").update({ Title: "ProjectType" });
		}
	} catch (error) {
		console.warn(`Failed to rename Title field on list ${LIST_TITLE}:`, error);
	}
}

function formatNumber(value: number): string {
	return Number.isFinite(value) ? `${value}` : "";
}

async function ensureSeedData(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const existingItems = await list.items.select(
		"Id",
		"Title",
		"Simple",
		"Medium",
		"Complex",
		"VeryComplex"
	)();

	const existingMap = new Map<
		string,
		{
			id: number;
			values: Record<keyof Omit<SeedItem, "LinkTitle">, string>;
		}
	>();

	for (const item of existingItems) {
		const title = `${item.Title ?? ""}`;
		if (!title) {
			continue;
		}
		existingMap.set(title, {
			id: Number(item.Id),
			values: {
				Simple: `${item.Simple ?? ""}`,
				Medium: `${item.Medium ?? ""}`,
				Complex: `${item.Complex ?? ""}`,
				VeryComplex: `${item.VeryComplex ?? ""}`
			}
		});
	}

	const fields: Array<keyof Omit<SeedItem, "LinkTitle">> = [
		"Simple",
		"Medium",
		"Complex",
		"VeryComplex"
	];

	for (const seed of seedItems) {
		const payload = {
			Title: seed.LinkTitle,
			Simple: formatNumber(seed.Simple),
			Medium: formatNumber(seed.Medium),
			Complex: formatNumber(seed.Complex),
			VeryComplex: formatNumber(seed.VeryComplex)
		};

		const existing = existingMap.get(seed.LinkTitle);
		if (!existing) {
			await list.items.add(payload);
			continue;
		}

		const updates: Partial<typeof payload> = {};
		for (const field of fields) {
			const target = payload[field];
			const current = existing.values[field];
			if (current !== target) {
				updates[field] = target;
			}
		}

		if (Object.keys(updates).length > 0) {
			await list.items.getById(existing.id).update(updates);
		}
	}
}

export async function provisionComplexityWeightage(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureTitleRenamed(sp);
	await ensureSeedData(sp);
}

export default provisionComplexityWeightage;
