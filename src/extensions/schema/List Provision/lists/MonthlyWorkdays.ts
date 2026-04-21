/*eslint-disable*/
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

const LIST_TITLE = RequiredListsProvision.MonthlyWorkdays;

type MonthlyWorkdaysFieldName = "Year" | "Days";
type MonthlyWorkdaysViewField = MonthlyWorkdaysFieldName;

const fieldDefinitions: readonly FieldDefinition<MonthlyWorkdaysFieldName>[] = [
	{
		internalName: "Year",
		schemaXml: "<Field Type='Number' Name='Year' StaticName='Year' DisplayName='Year' />"
	},
	{
		internalName: "Days",
		schemaXml: "<Field Type='Number' Name='Days' StaticName='Days' DisplayName='Days' />"
	}
] as const;

const defaultViewFields: readonly MonthlyWorkdaysViewField[] = [
	"Year",
	"Days"
] as const;

const definition: ListProvisionDefinition<MonthlyWorkdaysFieldName, MonthlyWorkdaysViewField> = {
	title: LIST_TITLE,
	description: "Monthly workdays list",
	templateId: 100,
	fields: fieldDefinitions,
	defaultViewFields
};

export async function provisionMonthlyWorkdays(sp: SPFI): Promise<void> {
	await ensureListProvision(sp, definition);
}

export default provisionMonthlyWorkdays;
