import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

export const RequiredListsProvision = {
    ProjectMetricLogs: "ProjectMetricLogs",
    EmailLogs: "EmailLogs",
    ProjectMetrics: "ProjectMetrics"
};

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricLogs');
    const { provisionEmailLogs } = await import('./lists/EmailLogs');
    const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');

    await provisionProjectMetricLogs(sp);
    await provisionEmailLogs(sp);
    await provisionProjectMetrics(sp);
}

export type FieldDefinition<TInternalName extends string> = {
    internalName: TInternalName;
    schemaXml: string;
};

interface ViewDefinition<TViewField extends string> {
    title: string;
    fields: readonly TViewField[];
    rowLimit?: number;
    makeDefault?: boolean;
}

export interface ListProvisionDefinition<TFieldName extends string, TViewField extends string = TFieldName> {
    title: string;
    description?: string;
    templateId?: number;
    fields?: readonly FieldDefinition<TFieldName>[];
    defaultViewFields?: readonly TViewField[];
    /** Optional list of field internal names to remove from the list if they exist */
    removeFields?: readonly TFieldName[];
    views?: readonly ViewDefinition<TViewField>[];
}

async function fieldExists<TFieldName extends string>(sp: SPFI, listTitle: string, fieldName: TFieldName): Promise<boolean> {
    try {
        const field = await sp.web.lists.getByTitle(listTitle).fields.getByInternalNameOrTitle(fieldName).select("Id")();
        return !!field;
    } catch (error) {
        return false;
    }
}

export async function ensureListProvision<TFieldName extends string, TViewField extends string = TFieldName>(
    sp: SPFI,
    definition: ListProvisionDefinition<TFieldName, TViewField>
): Promise<void> {
    const {
        title,
        description = "",
        templateId = 100,
        fields = [],
        defaultViewFields = [],
        removeFields = [],
        views = []
    } = definition;

    const ensureResult = await sp.web.lists.ensure(title, description, templateId);
    const list = ensureResult.list;

    for (const field of fields) {
        const exists = await fieldExists(sp, title, field.internalName);
        if (!exists) {
            await list.fields.createFieldAsXml(field.schemaXml);
        }
    }

    if (removeFields.length > 0) {
        for (const fieldName of removeFields) {
            try {
                const exists = await fieldExists(sp, title, fieldName);
                if (exists) {
                    await list.fields.getByInternalNameOrTitle(fieldName).delete();
                    const defaultView = list.defaultView;
                    const schemaXml = await defaultView.fields.getSchemaXml();
                    if (schemaXml.includes(`Name=\"${fieldName}\"`) || schemaXml.includes(`Name='${fieldName}'`)) {
                        await defaultView.fields.remove(fieldName);
                    }
                }
            } catch (err) {
                console.warn(`Failed to check existence of field ${fieldName} on list ${title}:`, err);
            }
        }
    }

    if (defaultViewFields.length > 0) {
        const defaultView = list.defaultView;
        const schemaXml = await defaultView.fields.getSchemaXml();
        for (const viewField of defaultViewFields) {
            if (!schemaXml.includes(`Name=\"${viewField}\"`) && !schemaXml.includes(`Name='${viewField}'`)) {
                await defaultView.fields.add(viewField);
            }
        }
    }

    if (views.length > 0) {
        for (const view of views) {
            await addViewToList(sp, title, view);
        }
    }
}

async function addViewToList<TViewField extends string>(
    sp: SPFI,
    listTitle: string,
    view: ViewDefinition<TViewField>
): Promise<void> {
    const list = sp.web.lists.getByTitle(listTitle);
    const existingViews = await list.views.select("Title")();
    const alreadyExists = existingViews.some((existing: { Title: string }) => existing.Title === view.title);
    if (alreadyExists) {
        return;
    }

    const viewFields = [...view.fields];
    await list.views.add(view.title, false, viewFields);

    const updates: Record<string, unknown> = {};
    if (view.rowLimit !== undefined) {
        updates.RowLimit = view.rowLimit;
    }
    if (view.makeDefault) {
        updates.SetAsDefaultView = true;
    }

    if (Object.keys(updates).length > 0) {
        await list.views.getByTitle(view.title).update(updates);
    }
}