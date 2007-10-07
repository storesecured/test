<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

sQuestion_Path = "advanced/custom_fields.htm"
sInstructions="Custom fields can be used to collect information at checkout that the system doesnt already ask for.  These fields are requested of the user at checkout time and added to the bottom of the invoice."


set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_custom_fields"
myStructure("TableWhere") = "Store_ID="&Store_Id
'myStructure("ColumnList") = "customfield_id,custom_field_name,custom_field_type,custom_field_values,field_required,view_order"
myStructure("ColumnList") = "customfield_id,custom_field_name,field_required,view_order"
myStructure("DefaultSort") = "custom_field_name"
myStructure("PrimaryKey") = "customfield_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "general"
myStructure("Title") = "Custom Fields"
myStructure("FullTitle") = "General > Custom Fields"
myStructure("CommonName") = "Custom Field"
myStructure("NewRecord") = "custom_fields.asp"
myStructure("EditRecord") = "custom_fields.asp"
myStructure("Heading:customfield_id") = "PK"
myStructure("Heading:custom_field_name") = "Field Name"
'myStructure("Heading:custom_field_type") = "Field Type"
'myStructure("Heading:custom_field_values") = "Field Value"
myStructure("Heading:field_required") = "Field Required"
myStructure("Heading:view_order") = "View Order"
'myStructure("Format:customfield_id") = "INT"
myStructure("Format:custom_field_name") = "STRING"
'myStructure("Format:custom_field_type") = "INT"
'myStructure("Format:custom_field_type") = "LOOKUP"
'myStructure("Lookup:custom_field_type") = "1:TF^2:TA"
'myStructure("Format:custom_field_values") = "STRING"
myStructure("Format:field_required") = "LOOKUP"
myStructure("Lookup:field_required") = "0:No^-1:Yes"
'myStructure("Format:custom_fIeld_type") = "INT"

myStructure("Format:view_order") = "INT"

sTextHelp="fields/custom_fields.doc"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%

Delete_Id = fn_get_delete_ids
if Delete_Id<>"" then
    sql_delete = "delete from Store_custom_fields where customfield_Id in("&delete_id&") and Store_id = "&Store_id
    conn_store.Execute sql_delete
end if

set myStructure=nothing
 createFoot thisRedirect, 0%>
