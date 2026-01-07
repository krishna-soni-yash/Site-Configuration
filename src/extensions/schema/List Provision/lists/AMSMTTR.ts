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

const LIST_TITLE = RequiredListsProvision.AMSMTTR;

type AMSMTTRFieldName =
	| "responseTimeMin"
	| "responseTimeMax"
	| "resolutionTimeMin"
	| "resolutionTimeMax";

type AMSMTTRViewField = AMSMTTRFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<AMSMTTRFieldName>[] = [
	{
		internalName: "responseTimeMin",
		schemaXml: `<Field Type='Number' Name='responseTimeMin' StaticName='responseTimeMin' DisplayName='responseTimeMin' Decimals='2' />`
	},
	{
		internalName: "responseTimeMax",
		schemaXml: `<Field Type='Number' Name='responseTimeMax' StaticName='responseTimeMax' DisplayName='responseTimeMax' Decimals='2' />`
	},
	{
		internalName: "resolutionTimeMin",
		schemaXml: `<Field Type='Number' Name='resolutionTimeMin' StaticName='resolutionTimeMin' DisplayName='resolutionTimeMin' Decimals='2' />`
	},
	{
		internalName: "resolutionTimeMax",
		schemaXml: `<Field Type='Number' Name='resolutionTimeMax' StaticName='resolutionTimeMax' DisplayName='resolutionTimeMax' Decimals='2' />`
	}
] as const;

const defaultViewFields: readonly AMSMTTRViewField[] = [
	"LinkTitle",
	"responseTimeMin",
	"responseTimeMax",
	"resolutionTimeMin",
	"resolutionTimeMax"
] as const;

const definition: ListProvisionDefinition<AMSMTTRFieldName, AMSMTTRViewField> = {
	title: LIST_TITLE,
	description: "AMSMTTR list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

type SeedItem = {
	Title: string;
	responseTimeMin: number;
	responseTimeMax: number;
	resolutionTimeMin: number;
	resolutionTimeMax: number;
};

const seedItems: readonly SeedItem[] = [
	{
		Title: "P1",
		responseTimeMin: 0.05,
		responseTimeMax: 0.5,
		resolutionTimeMin: 0.5,
		resolutionTimeMax: 6
	},
	{
		Title: "P2",
		responseTimeMin: 0.05,
		responseTimeMax: 1,
		resolutionTimeMin: 0.5,
		resolutionTimeMax: 12
	},
	{
		Title: "P3",
		responseTimeMin: 0.05,
		responseTimeMax: 5,
		resolutionTimeMin: 0.5,
		resolutionTimeMax: 37
	},
	{
		Title: "P4",
		responseTimeMin: 0.05,
		responseTimeMax: 7,
		resolutionTimeMin: 0.5,
		resolutionTimeMax: 68
	},
];

async function ensureTitleRenamed(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	try {
		const field = await list.fields.getByInternalNameOrTitle("Title").select("Title")();
		if (!field || `${field.Title}` !== "priority") {
			await list.fields.getByInternalNameOrTitle("Title").update({ Title: "priority" });
		}
	} catch (error) {
		console.warn(`Failed to rename Title field on list ${LIST_TITLE}:`, error);
	}
}

function toNumber(value: unknown): number {
	if (typeof value === "number") {
		return value;
	}
	const parsed = Number(`${value ?? ""}`);
	return Number.isFinite(parsed) ? parsed : 0;
}

async function ensureSeedData(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const existingItems = await list.items.select(
		"Id",
		"Title",
		"responseTimeMin",
		"responseTimeMax",
		"resolutionTimeMin",
		"resolutionTimeMax"
	)();

	const existingMap = new Map<
		string,
		{
			id: number;
			responseTimeMin: number;
			responseTimeMax: number;
			resolutionTimeMin: number;
			resolutionTimeMax: number;
		}
	>();

	for (const item of existingItems) {
		const title = `${item.Title ?? ""}`;
		if (!title) {
			continue;
		}
		existingMap.set(title, {
			id: Number(item.Id),
			responseTimeMin: toNumber(item.responseTimeMin),
			responseTimeMax: toNumber(item.responseTimeMax),
			resolutionTimeMin: toNumber(item.resolutionTimeMin),
			resolutionTimeMax: toNumber(item.resolutionTimeMax)
		});
	}

	const numericFields: Array<keyof Omit<SeedItem, "Title">> = [
		"responseTimeMin",
		"responseTimeMax",
		"resolutionTimeMin",
		"resolutionTimeMax"
	];

	for (const seed of seedItems) {
		const payload = {
			Title: seed.Title,
			responseTimeMin: seed.responseTimeMin,
			responseTimeMax: seed.responseTimeMax,
			resolutionTimeMin: seed.resolutionTimeMin,
			resolutionTimeMax: seed.resolutionTimeMax
		};

		const existing = existingMap.get(seed.Title);
		if (!existing) {
			await list.items.add(payload);
			continue;
		}

		const updates: Partial<typeof payload> = {};
		for (const field of numericFields) {
			const target = seed[field];
			const current = existing[field];
			if (Math.abs(current - target) > 0.0001) {
				updates[field] = target;
			}
		}

		if (Object.keys(updates).length > 0) {
			await list.items.getById(existing.id).update(updates);
		}
	}
}

export async function provisionAMSMTTR(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureTitleRenamed(sp);
	await ensureSeedData(sp);
}

export default provisionAMSMTTR;
