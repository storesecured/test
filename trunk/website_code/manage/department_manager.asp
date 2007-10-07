<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%



sQuestion_Path = "inventory/departments.htm"
sInstructions="Departments are used to group inventory together for easier browsing.	Departments may have as many levels of subdepartments as needed."
sFlashHelp="departments/departments.htm"
sMediaHelp="departments/departments.wmv"
sZipHelp="departments/departments.zip"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_dept"
myStructure("ColumnList") = "department_id,department_name,full_name"
myStructure("HeaderList") = "department_id,department_name,full_name,details"
myStructure("DefaultSort") = "full_name"
myStructure("TableWhere") = "department_id<>0"
myStructure("PrimaryKey") = "department_id"
myStructure("Footer") = "<input type=""button"" OnClick=JavaScript:self.location=""store_dept_add.asp"" class=""Buttons"" value=""Advanced Add Department"" name=""Create_new"">"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "inventory"
myStructure("FileName") = "department_manager.asp"
myStructure("FormAction") = "department_manager.asp"
myStructure("Title") = "Departments"
myStructure("FullTitle") = "Inventory > Departments"
myStructure("CommonName") = "Department"
myStructure("NewRecord") = "store_dept_basic.asp"
myStructure("EditRecord") = "store_dept_basic_edit.asp"
myStructure("Heading:details") = "Advanced Edit"
myStructure("Heading:department_id") = "Id"
myStructure("Heading:department_name") = "Name"
myStructure("Heading:full_name") = "Full"
myStructure("Format:details") = "TEXT"
myStructure("Format:department_id") = "INT"
myStructure("Format:department_name") = "STRING"
myStructure("Format:full_name") = "STRING"
myStructure("Link:details") = "store_dept_edit.asp?Id=PK"
%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%

Delete_Id = fn_get_delete_ids

if Delete_Id<>"" then
    sql_delete = "delete from Store_items_dept where store_Id="&Store_Id&" and department_id in ("&delete_id&")"	
	conn_store.Execute sql_delete
	
	server.execute "reset_design.asp"
	response.redirect "department_manager.asp"
end if

createFoot thisRedirect, 0
%>

