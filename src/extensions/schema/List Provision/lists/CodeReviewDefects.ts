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

const LIST_TITLE = RequiredListsProvision.CodeReviewDefects;

type CodeReviewDefectsFieldName =
    | "Code File Cla"
    | "Code File Cla0"
    | "Code File -"
    | "Reviewer Name"
    | "Identified Date"
    | "Defect Type"
    | "Defect Classification"
    | "Defect Description"
    | "Code Review Checklis"
    | "Severity"
    | "Defect Status"
    | "Review Results"
    | "Defect Origin Phase"
    | "Impacted Components"
    | "Correction / C"
    | "Planned Closure Date"
    | "Actual Closure Date"
    | "Review Iteration Num"
    | "Review Completion Da"
    | "Location of defect ("
    | "Remarks";
type CodeReviewDefectsViewField = CodeReviewDefectsFieldName;

function buildFieldDefinitions(): FieldDefinition<CodeReviewDefectsFieldName>[] {
    return [
        {
            internalName: "Code File Cla",
            schemaXml: `<Field Type='Text' Name='Code File Cla' StaticName='Code File Class' DisplayName='Code File/ Class Name' MaxLength='255' />`
        }
        ,
        {
            internalName: "Code File Cla0",
            schemaXml: `<Field Type='Text' Name='Code File Cla0' StaticName='Code File Cla0' DisplayName='Code File/ Class Size (Lines of Code)' MaxLength='255' />`
        }
        ,
        {
            internalName: "Code File -",
            schemaXml: `<Field Type='User' Name='Code File -' StaticName='Code File -' DisplayName='Code File - Author/ Developer Name(s)' UserSelectionMode='PeopleOnly' />`
        }
        ,
        {
            internalName: "Reviewer Name",
            schemaXml: `<Field Type='User' Name='Reviewer Name' StaticName='Reviewer Name' DisplayName='Reviewer Name' UserSelectionMode='PeopleOnly' />`
        }
        ,
        {
            internalName: "Identified Date",
            schemaXml: `<Field Type='DateTime' Name='Identified Date' StaticName='Identified Date' DisplayName='Identified Date' Format='DateOnly' />`
        }
        ,
        {
            internalName: "Defect Type",
            schemaXml: `<Field Type='Choice' Name='Defect Type' StaticName='Defect Type' DisplayName='Defect Type' Format='Dropdown'><CHOICES><CHOICE>Function</CHOICE><CHOICE>Interface</CHOICE><CHOICE>Checking</CHOICE><CHOICE>Assignment</CHOICE><CHOICE>Timing/ Serialization</CHOICE><CHOICE>Build/ Package</CHOICE><CHOICE>Documentation</CHOICE><CHOICE>Algorithm</CHOICE></CHOICES></Field>`
        }
        ,
        {
            internalName: "Defect Classification",
            schemaXml: `<Field Type='Choice' Name='Defect Classification' StaticName='Defect Classification' DisplayName='Defect Classification' Format='Dropdown'><CHOICES><CHOICE>Design issues</CHOICE><CHOICE>Data Validation</CHOICE><CHOICE>Logical</CHOICE><CHOICE>Computational</CHOICE><CHOICE>User Interface Presentation</CHOICE><CHOICE>Input/output</CHOICE><CHOICE>Error Handling/Exception Handling</CHOICE><CHOICE>Initialization</CHOICE><CHOICE>Installation/Configuration</CHOICE><CHOICE>Performance/Load/Stress</CHOICE><CHOICE>Online help / Error messages</CHOICE><CHOICE>Third Party Software</CHOICE><CHOICE>Not Meeting Coding Standard</CHOICE></CHOICES></Field>`
        }
        ,
        {
            internalName: "Defect Description",
            schemaXml: `<Field Type='Note' Name='Defect Description' StaticName='Defect Description' DisplayName='Defect Description' NumLines='6' RichText='FALSE' />`
        }
        ,
        {
            internalName: "Code Review Checklis",
            schemaXml: `<Field Type='Text' Name='Code Review Checklis' StaticName='Code Review Checklis' DisplayName='Code Review Checklist/ Coding Standards section ref. no.' MaxLength='255' />`
        }
        ,
        {
            internalName: "Severity",
            schemaXml: `<Field Type='Choice' Name='Severity' StaticName='Severity' DisplayName='Severity' Format='Dropdown'><CHOICES><CHOICE>Critical</CHOICE><CHOICE>Moderate</CHOICE><CHOICE>Minor</CHOICE></CHOICES></Field>`
        }
        ,
        {
            internalName: "Defect Status",
            schemaXml: `<Field Type='Choice' Name='Defect Status' StaticName='Defect Status' DisplayName='Defect Status' Format='Dropdown'><CHOICES><CHOICE>Open</CHOICE><CHOICE>Closed</CHOICE><CHOICE>Defered</CHOICE><CHOICE>Inprocess</CHOICE></CHOICES></Field>`
        }
        ,
        {
            internalName: "Review Results",
            schemaXml: `<Field Type='Choice' Name='Review Results' StaticName='Review Results' DisplayName='Review Results' Format='Dropdown'><CHOICES><CHOICE>Accepted</CHOICE><CHOICE>Rejected</CHOICE><CHOICE>Accepted with Re review</CHOICE><CHOICE>Rejected with Re review</CHOICE></CHOICES></Field>`
        }
        ,
        {
            internalName: "Defect Origin Phase",
            schemaXml: `<Field Type='Choice' Name='Defect Origin Phase' StaticName='Defect Origin Phase' DisplayName='Defect Origin Phase' Format='Dropdown'><CHOICES><CHOICE>Requirements Specification</CHOICE><CHOICE>Designs</CHOICE><CHOICE>Code</CHOICE><CHOICE>Unit Testing</CHOICE><CHOICE>Integration Testing</CHOICE><CHOICE>System Testing</CHOICE><CHOICE>Acceptance Testing</CHOICE><CHOICE>Others</CHOICE></CHOICES></Field>`
        }
        ,
        {
            internalName: "Impacted Components",
            schemaXml: `<Field Type='Text' Name='Impacted Components' StaticName='Impacted Components' DisplayName='Impacted Components' MaxLength='255' />`
        }
        ,
        {
            internalName: "Correction / C",
            schemaXml: `<Field Type='Note' Name='Correction / C' StaticName='Correction / C' DisplayName='Correction / Corrective Action' NumLines='6' RichText='FALSE' />`
        }
        ,
        {
            internalName: "Planned Closure Date",
            schemaXml: `<Field Type='DateTime' Name='Planned Closure Date' StaticName='Planned Closure Date' DisplayName='Planned Closure Date' Format='DateOnly' />`
        }
        ,
        {
            internalName: "Actual Closure Date",
            schemaXml: `<Field Type='DateTime' Name='Actual Closure Date' StaticName='Actual Closure Date' DisplayName='Actual Closure Date' Format='DateOnly' />`
        }
        ,
        {
            internalName: "Review Iteration Num",
            schemaXml: `<Field Type='Number' Name='Review Iteration Num' StaticName='Review Iteration Num' DisplayName='Review Iteration Number' />`
        }
        ,
        {
            internalName: "Review Completion Da",
            schemaXml: `<Field Type='DateTime' Name='Review Completion Da' StaticName='Review Completion Da' DisplayName='Review Completion Date' Format='DateOnly' />`
        }
        ,
        {
            internalName: "Location of defect (",
            schemaXml: `<Field Type='Text' Name='Location of defect (' StaticName='Location of defect (' DisplayName='Location of defect (Sub section - Line Number)' MaxLength='255' />`
        }
        ,
        {
            internalName: "Remarks",
            schemaXml: `<Field Type='Text' Name='Remarks' StaticName='Remarks' DisplayName='Remarks' MaxLength='255' />`
        }
    ];
}

const defaultViewFields: readonly CodeReviewDefectsViewField[] = [
    "Code File Cla",
    "Code File Cla0",
    "Code File -",
    "Reviewer Name",
    "Identified Date",
    "Defect Type",
    "Defect Classification",
    "Defect Description",
    "Code Review Checklis",
    "Severity",
    "Defect Status",
    "Review Results",
    "Defect Origin Phase",
    "Impacted Components",
    "Correction / C",
    "Planned Closure Date",
    "Actual Closure Date",
    "Review Iteration Num",
    "Review Completion Da",
    "Location of defect (",
    "Remarks"
] as const;

const viewDefinitions: ReadonlyArray<{
    title: string;
    fields: readonly CodeReviewDefectsViewField[];
    makeDefault?: boolean;
    includeLinkTitle?: boolean;
}> = [
    {
        title: "All Code Review Defects",
        fields: defaultViewFields,
        makeDefault: true,
        includeLinkTitle: true
    }
];

async function ensureDefaultViewIncludesFields(
    list: any,
    fields: readonly CodeReviewDefectsViewField[]
): Promise<void> {
    const defaultView = list.defaultView;
    let schemaXml = "";

    try {
        schemaXml = await defaultView.fields.getSchemaXml();
    } catch (error) {
        console.warn(`Failed to read default view schema for list ${LIST_TITLE}:`, error);
        return;
    }

    for (const field of fields) {
        const singleQuoteToken = `Name='${field}'`;
        const doubleQuoteToken = `Name="${field}"`;
        if (schemaXml.includes(singleQuoteToken) || schemaXml.includes(doubleQuoteToken)) {
            continue;
        }

        try {
            await defaultView.fields.add(field as any);
            schemaXml += singleQuoteToken;
        } catch (error) {
            console.warn(`Failed to add field ${field} to default view for list ${LIST_TITLE}:`, error);
        }
    }
}

export async function provisionCodeReviewDefects(sp: SPFI): Promise<void> {
    const fields = buildFieldDefinitions();

    const definition: ListProvisionDefinition<
        CodeReviewDefectsFieldName,
        CodeReviewDefectsViewField
    > = {
        title: LIST_TITLE,
        description: "Code review defects list",
        templateId: 100,
        fields,
        defaultViewFields,
        views: viewDefinitions
    };

    await ensureListProvision(sp, definition);
    const list = sp.web.lists.getByTitle(LIST_TITLE);
    await list.fields.getByInternalNameOrTitle("Title").update({ Title: "Requirement ID / Ticket ID" });
    await ensureDefaultViewIncludesFields(list, defaultViewFields);
}

export default provisionCodeReviewDefects;
