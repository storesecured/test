<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

Delete_Id = fn_get_delete_ids
if Delete_Id<>"" then
	sql_check= "select template_id from store_design_template where store_id="&Store_Id&" and template_id in ("&delete_id&") and (last_selected=1 or template_active=1)"
	set rs_check=conn_store.execute(sql_check)
	if rs_check.EOF then
		sql_delete = "delete from Store_design_objects where template_Id in ("&delete_id&") and Store_id = "&Store_id
		conn_store.Execute sql_delete
	else
	   fn_error "A template that is currently applied to your store cannot be deleted"
	end if
end if

sInstructions="Custom templates can be used to manage and setup different design templates for your store."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_design_template"
myStructure("TableWhere") = "template_active=0 and Store_ID="&Store_Id
myStructure("ColumnList") = "template_id,template_name,last_selected,template_description"
myStructure("DefaultSort") = "template_name"
myStructure("PrimaryKey") = "template_id"
myStructure("Level") = 1
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "design"
myStructure("Title") = "Templates"
myStructure("FullTitle") = "Design > Templates"
myStructure("CommonName") = "Template"
myStructure("NewRecord") = "custom_template.asp"
myStructure("EditRecord") = "layout_design.asp"
myStructure("Heading:template_name") = "Name"
myStructure("Heading:template_id") = "PK"
myStructure("Heading:last_selected") = "Last Selected"
myStructure("Heading:template_description") = "Description"

myStructure("Format:template_name") = "STRING"
myStructure("Format:last_selected") = "LOOKUP"
myStructure("Lookup:last_selected") = "0:No^-1:Yes"

myStructure("Format:template_description") = "STRING"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->


<%
set myStructure=nothing
 createFoot thisRedirect, 0%>
