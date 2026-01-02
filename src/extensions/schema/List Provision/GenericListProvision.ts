/*eslint-disable*/
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/content-types";
import "@pnp/sp/folders";

export type FieldDefinition<TInternalName extends string> = {
    internalName: TInternalName;
    schemaXml: string;
};

interface ViewDefinition<TViewField extends string> {
    title: string;
    fields: readonly TViewField[];
    rowLimit?: number;
    makeDefault?: boolean;
    includeLinkTitle?: boolean;
    query?: string;
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

export interface ContentTypeBindingDefinition {
    id: string;
    name?: string;
    group?: string;
    ensureOnWeb?: boolean;
}

export interface EnsureListContentTypeOptions {
    ensureOnWeb?: boolean;
    order?: readonly string[];
    removeDefaultContentType?: boolean;
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
                await defaultView.fields.add(viewField as any);
            }
        }
    }

    if (views.length > 0) {
        for (const view of views) {
            await addViewToList(sp, title, view as ViewDefinition<string>);
        }
    }
}

export async function addViewToList<TViewField extends string>(
    sp: SPFI,
    listTitle: string,
    view: ViewDefinition<TViewField>
): Promise<void> {
    const list = sp.web.lists.getByTitle(listTitle);
    const includeLinkTitle = view.includeLinkTitle ?? true;
    const viewFields = [...(view.fields as readonly string[])];
    if (includeLinkTitle && !viewFields.includes("LinkTitle")) {
        viewFields.unshift("LinkTitle");
    }

    const existingViews = await list.views.select("Id", "Title")();
    const existingView = existingViews.find(
        (existing: { Id: string; Title: string }) => existing.Title === view.title
    );

    const ensureView = async () => {
        if (existingView) {
            return list.views.getById(existingView.Id);
        }

        const added = await (list.views as any).add(view.title, false);
        if (added?.view) {
            return added.view as any;
        }
        if (added?.data?.Id) {
            return list.views.getById(added.data.Id);
        }
        return list.views.getByTitle(view.title);
    };

    const listView = await ensureView();

    try {
        if (typeof listView.fields?.removeAll === "function") {
            await listView.fields.removeAll();
        }
    } catch (error) {
        console.warn(`Failed to clear view fields for ${view.title} on ${listTitle}:`, error);
    }

    const appliedFields: string[] = [];
    for (const field of viewFields) {
        try {
            await listView.fields.add(field as any);
            appliedFields.push(field);
        } catch (error) {
            console.warn(`Failed to add field ${field} to view ${view.title} on list ${listTitle}:`, error);
        }
    }

    if (appliedFields.length === 0 && includeLinkTitle) {
        try {
            await listView.fields.add("LinkTitle" as any);
        } catch (error) {
            console.warn(`Failed to ensure LinkTitle on view ${view.title} for list ${listTitle}:`, error);
        }
    }

    const updates: Record<string, unknown> = {};
    if (view.rowLimit !== undefined) {
        updates.RowLimit = view.rowLimit;
    }

    if (view.query !== undefined) {
        updates.ViewQuery = view.query;
    }

    if (Object.keys(updates).length > 0) {
        await listView.update(updates);
    }

    if (view.makeDefault) {
        try {
            await listView.setAsDefault();
        } catch (error) {
            console.warn(`Failed to set ${view.title} as default view for list ${listTitle}:`, error);
        }
    }
}

// Generics methods to provison list using Content Types

export async function contentTypeExists(sp: SPFI, contentTypeId: string): Promise<boolean> {
    try {
        await sp.web.contentTypes.getById(contentTypeId).select("Id")();
        return true;
    } catch (error) {
        return false;
    }
}

async function bindContentTypeToList(
    sp: SPFI,
    list: any,
    listTitle: string,
    contentTypeId: string,
    contentTypeName?: string
): Promise<boolean> {
    let added = false;
    const listContentTypes = list?.contentTypes as any;

    if (listContentTypes && typeof listContentTypes.addAvailableContentType === "function") {
        try {
            await listContentTypes.addAvailableContentType(contentTypeId);
            added = true;
        } catch (error) {
            console.warn(`addAvailableContentType failed for ${contentTypeId} on list ${listTitle}:`, error);
        }
    }

    if (!added && listContentTypes && typeof listContentTypes.addExistingContentType === "function") {
        try {
            await listContentTypes.addExistingContentType(contentTypeId);
            added = true;
        } catch (error) {
            console.warn(`addExistingContentType failed for ${contentTypeId} on list ${listTitle}:`, error);
        }
    }

    if (!added && listContentTypes && typeof listContentTypes.add === "function") {
        try {
            await listContentTypes.add({ Id: { StringValue: contentTypeId }, Name: contentTypeName } as any);
            added = true;
        } catch (error) {
            console.warn(`list.contentTypes.add(...) failed for ${contentTypeId} on list ${listTitle}:`, error);
        }
    }

    if (!added) {
        try {
            await sp.web.lists.getByTitle(listTitle).contentTypes.addAvailableContentType(contentTypeId);
            added = true;
        } catch (error) {
            console.warn(
                `Fallback sp.web.lists.getByTitle(...).contentTypes.addAvailableContentType failed for ${contentTypeId} on list ${listTitle}:`,
                error
            );
        }
    }

    return added;
}

export async function ensureListContentTypes(
    sp: SPFI,
    listTitle: string,
    contentTypes: readonly ContentTypeBindingDefinition[],
    options: EnsureListContentTypeOptions = {}
): Promise<void> {
    const list = sp.web.lists.getByTitle(listTitle);

    try {
        await list.update({ ContentTypesEnabled: true });
    } catch (error) {
        // Some list templates do not support toggling content types.
    }

    for (const definition of contentTypes) {
        const ensureAtWeb = definition.ensureOnWeb ?? options.ensureOnWeb ?? false;
        if (ensureAtWeb && definition.name) {
            const exists = await contentTypeExists(sp, definition.id);
            if (!exists) {
                try {
                    await (sp.web.contentTypes as any).add(
                        definition.name,
                        definition.id,
                        definition.group || "Custom"
                    );
                } catch (error) {
                    console.warn(`Failed to create content type ${definition.id} (${definition.name}) at web scope:`, error);
                }
            }
        }

        try {
            await bindContentTypeToList(sp, list, listTitle, definition.id, definition.name);
        } catch (error) {
            console.warn(`Error while adding content type ${definition.id} to list ${listTitle}:`, error);
        }
    }

    const order = options.order ?? contentTypes.map((ct) => ct.id);
    if (order.length > 0) {
        try {
            const rootFolder = (list as any).rootFolder as any;
            await rootFolder.update({ ContentTypeOrder: order.map((id) => ({ StringValue: id })) });
        } catch (error) {
            // Ordering content types is best-effort.
        }
    }

    if (options.removeDefaultContentType) {
        try {
            const existing = await (list as any).contentTypes.select("Id", "Name")();
            const unwrapId = (idField: any): string => {
                if (!idField) { return ""; }
                if (typeof idField === "string") { return idField; }
                if (typeof idField === "object" && idField.StringValue) { return idField.StringValue; }
                return "";
            };

            const itemCt = (existing || []).find((ct: any) => {
                try {
                    const name = ct?.Name ? String(ct.Name) : "";
                    const idVal = unwrapId(ct?.Id).toLowerCase();
                    return name === "Item" || idVal.indexOf("0x01") === 0;
                } catch (error) {
                    return false;
                }
            });

            const itemId = unwrapId(itemCt?.Id);
            if (itemId) {
                try {
                    await (list as any).contentTypes.getById(itemId).delete();
                } catch (error) {
                    // Ignore failures when Item content type is locked or in use.
                }
            }
        } catch (error) {
            // Reading content types is best-effort.
        }
    }
}