import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    ensureListProvision,
    FieldDefinition,
    ListProvisionDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";
import { provisionApplicableGraphs } from "./ApplicableGraphs";

const LIST_TITLE = RequiredListsProvision.OpenIssues;

type OpenIssuesFieldName =
    | "ApplicableGraphID"
    | "OpenIssuesCount";

type OpenIssuesViewField = OpenIssuesFieldName;

function normalizeListId(id: unknown): string {
    const value = `${id ?? ""}`;
    if (value.length === 0) {
        throw new Error("Unable to resolve ApplicableGraphs list identifier.");
    }
    return value.startsWith("{") ? value : `{${value}}`;
}

async function resolveApplicableGraphsListId(sp: SPFI): Promise<string> {
    try {
        const listInfo = await sp.web.lists.getByTitle(RequiredListsProvision.ApplicableGraphs).select("Id")();
        return normalizeListId(listInfo.Id);
    } catch (error) {
        await provisionApplicableGraphs(sp);
        const ensuredInfo = await sp.web.lists.getByTitle(RequiredListsProvision.ApplicableGraphs).select("Id")();
        return normalizeListId(ensuredInfo.Id);
    }
}

function buildFieldDefinitions(applicableGraphsListId: string): FieldDefinition<OpenIssuesFieldName>[] {
    return [
        {
            internalName: "ApplicableGraphID",
            schemaXml: `<Field Type='Lookup' Name='ApplicableGraphID' StaticName='ApplicableGraphID' DisplayName='ApplicableGraphID' List='${applicableGraphsListId}' ShowField='ID' LookupId='TRUE' />`
        },
        {
            internalName: "OpenIssuesCount",
            schemaXml: `<Field Type='Number' Name='OpenIssuesCount' StaticName='OpenIssuesCount' DisplayName='OpenIssuesCount' Decimals='2' />`
        }
    ];
}

const defaultViewFields: readonly OpenIssuesViewField[] = [
    "ApplicableGraphID",
    "OpenIssuesCount",
] as const;

const definition: ListProvisionDefinition<OpenIssuesFieldName, OpenIssuesViewField> = {
    title: LIST_TITLE,
    description: "Open Issues",
    templateId: 100,
    fields: undefined,
    defaultViewFields
};

export async function provisionOpenIssues(sp: SPFI): Promise<void> {
    const applicableGraphsListId = await resolveApplicableGraphsListId(sp);
    const fields = buildFieldDefinitions(applicableGraphsListId);

    await ensureListProvision(sp, {
        ...definition,
        fields
    });
}

export default provisionOpenIssues;
