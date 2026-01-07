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

const LIST_TITLE = RequiredListsProvision.SDLCParams;

type SDLCParamsFieldName = "MinValue" | "MaxValue" | "SDLCLifeCycleStage" | "Order0";

type SDLCParamsViewField = SDLCParamsFieldName | "LinkTitle";

const fieldDefinitions: readonly FieldDefinition<SDLCParamsFieldName>[] = [
	{
		internalName: "MinValue",
		schemaXml: `<Field Type='Text' Name='MinValue' StaticName='MinValue' DisplayName='MinValue' MaxLength='255' />`
	},
	{
		internalName: "MaxValue",
		schemaXml: `<Field Type='Text' Name='MaxValue' StaticName='MaxValue' DisplayName='MaxValue' MaxLength='255' />`
	},
	{
		internalName: "SDLCLifeCycleStage",
		schemaXml: `<Field Type='Text' Name='SDLCLifeCycleStage' StaticName='SDLCLifeCycleStage' DisplayName='SDLCLifeCycleStage' MaxLength='255' />`
	},
	{
		internalName: "Order0",
		schemaXml: `<Field Type='Text' Name='Order0' StaticName='Order0' DisplayName='Order' MaxLength='255' />`
	}
] as const;

const defaultViewFields: readonly SDLCParamsViewField[] = [
	"LinkTitle",
	"MinValue",
	"MaxValue",
	"SDLCLifeCycleStage",
	"Order0"
] as const;

const definition: ListProvisionDefinition<SDLCParamsFieldName, SDLCParamsViewField> = {
	title: LIST_TITLE,
	description: "SDLC parameters list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

const seedItems: ReadonlyArray<{
	Title: string;
	SDLCLifeCycleStage: string;
	MinValue: string;
	MaxValue: string;
	Order0: string;
}> = [
	{
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Burndown Effort",
		MinValue: "6",
		MaxValue: "10.75",
		Order0: "1"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Requirement Analysis Effort",
		MinValue: "0.5",
		MaxValue: "2",
		Order0: "2"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Coding Effort",
		MinValue: "1.5",
		MaxValue: "4.5",
		Order0: "3"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Code Review Effort",
		MinValue: "0.1",
		MaxValue: "0.5",
		Order0: "4"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Code Review Rework Effort",
		MinValue: "0",
		MaxValue: "0.1",
		Order0: "5"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Unit Testing Effort",
		MinValue: "0.5",
		MaxValue: "0.75",
		Order0: "6"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Test Execution Effort",
		MinValue: "0.5",
		MaxValue: "3",
		Order0: "7"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Code Review Defects",
		MinValue: "0",
		MaxValue: "0.5",
		Order0: "8"
	},
    {
		Title: "Agile",
		SDLCLifeCycleStage: "Predicted Test Defects",
		MinValue: "0",
		MaxValue: "0.5",
		Order0: "9"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Burndown Effort",
		MinValue: "6",
		MaxValue: "18",
		Order0: "1"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Requirement Analysis Effort",
		MinValue: "0.5",
		MaxValue: "3",
		Order0: "2"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Coding Effort",
		MinValue: "2",
		MaxValue: "10",
		Order0: "3"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Code Review Effort",
		MinValue: "0.25",
		MaxValue: "1",
		Order0: "4"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Code Review Rework Effort",
		MinValue: "0",
		MaxValue: "0.5",
		Order0: "5"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Unit Testing Effort",
		MinValue: "0.5",
		MaxValue: "2",
		Order0: "6"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Test Execution Effort",
		MinValue: "0.25",
		MaxValue: "2",
		Order0: "7"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Code Review Defects",
		MinValue: "0",
		MaxValue: "0.5",
		Order0: "8"
	},
    {
		Title: "DEV",
		SDLCLifeCycleStage: "Predicted Test Defects",
		MinValue: "0",
		MaxValue: "0.5",
		Order0: "9"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Burndown Effort",
		MinValue: "5",
		MaxValue: "9",
		Order0: "1"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Requirement Analysis Effort",
		MinValue: "0.5",
		MaxValue: "2.5",
		Order0: "2"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Coding Effort",
		MinValue: "1.5",
		MaxValue: "4",
		Order0: "3"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Code Review Effort",
		MinValue: "0.1",
		MaxValue: "1",
		Order0: "4"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Code Review Rework Effort",
		MinValue: "0",
		MaxValue: "0.5",
		Order0: "5"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Unit Testing Effort",
		MinValue: "0.5",
		MaxValue: "2",
		Order0: "6"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Test Execution Effort",
		MinValue: "0.5",
		MaxValue: "2",
		Order0: "7"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Code Review Defects",
		MinValue: "0",
		MaxValue: "0.25",
		Order0: "8"
	},
    {
		Title: "DEVM",
		SDLCLifeCycleStage: "Predicted Test Defects",
		MinValue: "0",
		MaxValue: "0.25",
		Order0: "9"
	},
];

async function ensureSeedData(sp: SPFI): Promise<void> {
	const list = sp.web.lists.getByTitle(LIST_TITLE);
	const existing = await list.items.select(
		"Id",
		"Title",
		"SDLCLifeCycleStage",
		"MinValue",
		"MaxValue",
		"Order0"
	)();
	const existingMap = new Map<string, {
		id: number;
		SDLCLifeCycleStage: string;
		MinValue: string;
		MaxValue: string;
		Order0: string;
	}>();

	for (const item of existing) {
		if (typeof item.Title === "string" && item.Title.length > 0) {
			existingMap.set(item.Title, {
				id: Number(item.Id),
				SDLCLifeCycleStage: `${item.SDLCLifeCycleStage ?? ""}`,
				MinValue: `${item.MinValue ?? ""}`,
				MaxValue: `${item.MaxValue ?? ""}`,
				Order0: `${item.Order0 ?? ""}`
			});
		}
	}

	for (const seed of seedItems) {
		const match = existingMap.get(seed.Title);
		if (!match) {
			await list.items.add({
				Title: seed.Title,
				SDLCLifeCycleStage: seed.SDLCLifeCycleStage,
				MinValue: seed.MinValue,
				MaxValue: seed.MaxValue,
				Order0: seed.Order0
			});
			continue;
		}

		const needsUpdate =
			match.SDLCLifeCycleStage !== seed.SDLCLifeCycleStage ||
			match.MinValue !== seed.MinValue ||
			match.MaxValue !== seed.MaxValue ||
			match.Order0 !== seed.Order0;

		if (needsUpdate) {
			await list.items.getById(match.id).update({
				SDLCLifeCycleStage: seed.SDLCLifeCycleStage,
				MinValue: seed.MinValue,
				MaxValue: seed.MaxValue,
				Order0: seed.Order0
			});
		}
	}
}

export async function provisionSDLCParams(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
	await ensureSeedData(sp);
}

export default provisionSDLCParams;
