/*eslint-disable*/
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import {
    ensureListProvision,
    ListProvisionDefinition,
    FieldDefinition
} from "../GenericListProvision";
import { RequiredListsProvision } from "../RequiredListProvision";

const LIST_TITLE = RequiredListsProvision.ResourceUtilization;

const resourceUtilizationFieldNames = [
    "Goal",
    "USL",
    "LSL",
    "ProjectType",
    "AllocatedEffort",
    "ProjectEffort",
    "ResourceUtilization"
] as const;

type ResourceUtilizationFieldName = typeof resourceUtilizationFieldNames[number];
type ResourceUtilizationViewField = ResourceUtilizationFieldName | "LinkTitle";

const textFieldNames: readonly ResourceUtilizationFieldName[] = [
    "Goal",
    "ProjectType"
] as const;

const numberFieldNames: readonly ResourceUtilizationFieldName[] = [
    "USL",
    "LSL",
    "AllocatedEffort",
    "ProjectEffort",
    "ResourceUtilization"
] as const;

const defaultViewFields: readonly ResourceUtilizationViewField[] = [
    "LinkTitle",
    "Goal",
    "USL",
    "LSL",
    "ProjectType",
    "AllocatedEffort",
    "ProjectEffort",
    "ResourceUtilization"
] as const;

function buildFieldDefinitions(): FieldDefinition<ResourceUtilizationFieldName>[] {
    const definitions: FieldDefinition<ResourceUtilizationFieldName>[] = [];

    for (const internalName of textFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Text' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' MaxLength='255' />`
        });
    }

    for (const internalName of numberFieldNames) {
        definitions.push({
            internalName,
            schemaXml: `<Field Type='Number' Name='${internalName}' StaticName='${internalName}' DisplayName='${internalName}' />`
        });
    }

    return definitions;
}

export async function provisionResourceUtilization(sp: SPFI): Promise<void> {
    const fields = buildFieldDefinitions();

    const definition: ListProvisionDefinition<ResourceUtilizationFieldName, ResourceUtilizationViewField> = {
        title: LIST_TITLE,
        description: "Resource utilization list",
        templateId: 100,
        fields,
        defaultViewFields
    };

    await ensureListProvision(sp, definition);
}

export default provisionResourceUtilization;
