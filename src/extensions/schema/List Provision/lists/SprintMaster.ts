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

const LIST_TITLE = RequiredListsProvision.SprintMaster;

type SprintMasterFieldName =
    | "SprintStartDate"
    | "SprintEndDate"
    | "DurationInWeeks";

type SprintMasterFieldNameViewField = SprintMasterFieldName;

const fieldDefinitions: readonly FieldDefinition<SprintMasterFieldName>[] = [
    {
        internalName: "SprintStartDate",
        schemaXml: `<Field Type='DateTime' Name='SprintStartDate' StaticName='SprintStartDate' DisplayName='SprintStartDate' Format='DateOnly' />`
    },
    {
        internalName: "SprintEndDate",
        schemaXml: `<Field Type='DateTime' Name='SprintEndDate' StaticName='SprintEndDate' DisplayName='SprintEndDate' Format='DateOnly' />`
    },
    {
        internalName: "DurationInWeeks",
        schemaXml: `<Field Type='Number' Name='DurationInWeeks' StaticName='DurationInWeeks' DisplayName='DurationInWeeks' Decimals='2' />`
    }
] as const;

const defaultViewFields: readonly SprintMasterFieldNameViewField[] = [
    "SprintStartDate",
    "SprintEndDate",
    "DurationInWeeks"
] as const;

const definition: ListProvisionDefinition<SprintMasterFieldName, SprintMasterFieldNameViewField> = {
    title: LIST_TITLE,
    description: "",
    templateId: 100,
    fields: fieldDefinitions,
    defaultViewFields
};

export async function provisionSprintMaster(sp: SPFI): Promise<void> {
    await ensureListProvision(sp, definition);
}

export default provisionSprintMaster;