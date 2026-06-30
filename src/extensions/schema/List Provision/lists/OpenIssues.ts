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
import { fetchListId } from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.OpenIssues;

type OpenIssuesFieldName =
    | "ApplicableGraphID"
    | "OpenIssuesCount";

type OpenIssuesViewField = OpenIssuesFieldName;

async function resolveApplicableGraphsListId(sp: SPFI): Promise<string> {
    return fetchListId(sp, RequiredListsProvision.ApplicableGraphs);
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

export async function provisionOpenIssues(sp: SPFI, applicableGraphsListId?: string): Promise<void> {
    const resolvedApplicableGraphsListId = applicableGraphsListId ?? await resolveApplicableGraphsListId(sp);
    const fields = buildFieldDefinitions(resolvedApplicableGraphsListId);

    await ensureListProvision(sp, {
        ...definition,
        fields
    });
}

export default provisionOpenIssues;
