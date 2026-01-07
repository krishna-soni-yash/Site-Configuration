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

const LIST_TITLE = RequiredListsProvision.AdjustmentFactorValue;

type AdjustmentFactorValueFieldName =
	| "ProjectType"
	| "VeryHigh"
	| "High"
	| "Medium"
	| "Low"
	| "VeryLow"
	| "None"
	| "NoReuse";

type AdjustmentFactorValueViewField = AdjustmentFactorValueFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<AdjustmentFactorValueFieldName>[] = [
	{
		internalName: "ProjectType",
		schemaXml: `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`
	},
	{
		internalName: "VeryHigh",
		schemaXml: `<Field Type='Text' Name='VeryHigh' StaticName='VeryHigh' DisplayName='VeryHigh' MaxLength='255' />`
	},
	{
		internalName: "High",
		schemaXml: `<Field Type='Text' Name='High' StaticName='High' DisplayName='High' MaxLength='255' />`
	},
	{
		internalName: "Medium",
		schemaXml: `<Field Type='Text' Name='Medium' StaticName='Medium' DisplayName='Medium' MaxLength='255' />`
	},
	{
		internalName: "Low",
		schemaXml: `<Field Type='Text' Name='Low' StaticName='Low' DisplayName='Low' MaxLength='255' />`
	},
	{
		internalName: "VeryLow",
		schemaXml: `<Field Type='Text' Name='VeryLow' StaticName='VeryLow' DisplayName='VeryLow' MaxLength='255' />`
	},
	{
		internalName: "None",
		schemaXml: `<Field Type='Text' Name='None' StaticName='None' DisplayName='None' MaxLength='255' />`
	},
	{
		internalName: "NoReuse",
		schemaXml: `<Field Type='Text' Name='NoReuse' StaticName='NoReuse' DisplayName='NoReuse' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly AdjustmentFactorValueViewField[] = [
	"LinkTitle",
	"ProjectType",
	"VeryHigh",
	"High",
	"Medium",
	"Low",
	"VeryLow",
	"None",
	"NoReuse"
] as const;

const definition: ListProvisionDefinition<AdjustmentFactorValueFieldName, AdjustmentFactorValueViewField> = {
	title: LIST_TITLE,
	description: "Adjustment factor value list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

type SeedItem = {
	Title: string;
	ProjectType: string;
	VeryHigh: string;
	High: string;
	Medium: string;
	Low: string;
	VeryLow: string;
	None: string;
	NoReuse?: string;
};

const seedItems: readonly SeedItem[] = [
	{
		Title: "Application and Environment Adjustment Factor",
		ProjectType: "Dev",
		VeryHigh: "1.35",
		High: "1.21",
		Medium: "1.07",
		Low: "0.93",
		VeryLow: "0.79",
		None: "0.65",
		NoReuse: ""
	},
	{
		Title: "Skill Adjustment Factor",
		ProjectType: "Dev",
		VeryHigh: "",
		High: "0.9",
		Medium: "1.05",
		Low: "1.25",
		VeryLow: "",
		None: "",
		NoReuse: ""
	},
	{
		Title: "Reusability of Design and Code",
		ProjectType: "Dev",
		VeryHigh: "",
		High: "0.91",
		Medium: "0.94",
		Low: "0.97",
		VeryLow: "",
		None: "",
		NoReuse: "1"
	},
	{
		Title: "Extent of Automation in Project",
		ProjectType: "Dev",
		VeryHigh: "",
		High: "0.9",
		Medium: "0.94",
		Low: "0.97",
		VeryLow: "",
		None: "1",
		NoReuse: ""
	}
] as const;

async function ensureSeedData(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const existingItems = await list.items.select(
		"Id",
		"Title",
		"ProjectType",
		"VeryHigh",
		"High",
		"Medium",
		"Low",
		"VeryLow",
		"None",
		"NoReuse"
	)();

	const existingMap = new Map<string, { id: number; values: Record<string, string> }>();
	for (const item of existingItems) {
		const title = `${item.Title ?? ""}`;
		const projectType = `${item.ProjectType ?? ""}`;
		if (!title) {
			continue;
		}
		const key = `${title}||${projectType}`;
		existingMap.set(key, {
			id: Number(item.Id),
			values: {
				ProjectType: projectType,
				VeryHigh: `${item.VeryHigh ?? ""}`,
				High: `${item.High ?? ""}`,
				Medium: `${item.Medium ?? ""}`,
				Low: `${item.Low ?? ""}`,
				VeryLow: `${item.VeryLow ?? ""}`,
				None: `${item.None ?? ""}`,
				NoReuse: `${item.NoReuse ?? ""}`
			}
		});
	}

	const fieldsToSync: Array<keyof SeedItem> = [
		"ProjectType",
		"VeryHigh",
		"High",
		"Medium",
		"Low",
		"VeryLow",
		"None",
		"NoReuse"
	];

	for (const seed of seedItems) {
		const key = `${seed.Title}||${seed.ProjectType}`;
		const payload: Record<string, string> = {
			Title: seed.Title,
			ProjectType: seed.ProjectType,
			VeryHigh: seed.VeryHigh,
			High: seed.High,
			Medium: seed.Medium,
			Low: seed.Low,
			VeryLow: seed.VeryLow,
			None: seed.None,
			NoReuse: seed.NoReuse ?? ""
		};

		const existing = existingMap.get(key);
		if (!existing) {
			await list.items.add(payload);
			continue;
		}

		const updates: Record<string, string> = {};
		for (const field of fieldsToSync) {
			const target = payload[field] ?? "";
			const current = existing.values[field as keyof typeof existing.values] ?? "";
			if (current !== target) {
				updates[field] = target;
			}
		}

		if (Object.keys(updates).length > 0) {
			await list.items.getById(existing.id).update(updates);
		}
	}
}

async function ensureTitleRenamed(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	try {
		const field = await list.fields.getByInternalNameOrTitle("Title").select("Title")();
		if (!field || `${field.Title}` !== "FactorName") {
			await list.fields.getByInternalNameOrTitle("Title").update({ Title: "FactorName" });
		}
	} catch (error) {
		console.warn(`Failed to rename Title field on list ${LIST_TITLE}:`, error);
	}
}

export async function provisionAdjustmentFactorValue(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureTitleRenamed(sp);
	await ensureSeedData(sp);
}

export default provisionAdjustmentFactorValue;
