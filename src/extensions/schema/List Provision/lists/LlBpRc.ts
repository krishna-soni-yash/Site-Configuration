import { SPFI } from "@pnp/sp";
import { RequiredListsProvision } from "../RequiredListProvision";
import { ensureListProvision, ensureListContentTypes, ContentTypeBindingDefinition } from "../GenericListProvision";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/content-types";
import "@pnp/sp/fields";
import "@pnp/sp/views";

const LIST_TITLE = RequiredListsProvision.LlBpRc;

const CONTENT_TYPES: ReadonlyArray<ContentTypeBindingDefinition> = [
    { id: "0x0100675A5D902917F5468E54C67BEC4A6765", name: "Lessons Learnt" },
    { id: "0x010013D4E57092D07541AB01F187F1C1A283", name: "Best Practices" },
    { id: "0x010063DA7C6C73EA594FA193E1333337F0E1", name: "Reusable Component" }
];

export async function provisionLlBpRc(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, {
		title: LIST_TITLE,
		description: "Lessons Learnt, Best Practices, Reusable Component list",
		templateId: 100,
		defaultViewFields: ["LinkTitle"],
		// views: [
		// 	{
		// 		title: "Lessons Learnt",
		// 		fields: ["ProblemFacedLearning", "Category", "Solution", "Remarks"],
		// 		makeDefault: true
		// 	},
        //     {
		// 		title: "Best Practices",
		// 		fields: ["BestPracticeDescription", "Reference", "Responsibility", "Remarks"]
		// 	},
        //     {
		// 		title: "Reusable Component",
		// 		fields: ["ComponentName", "Location", "PurposeMainFunctionality", "Responsibility", "Remarks"]
		// 	}
		// ]
	});

	await ensureListContentTypes(sp, LIST_TITLE, CONTENT_TYPES, {
		ensureOnWeb: true,
		removeDefaultContentType: true
	});
}

export default provisionLlBpRc;
