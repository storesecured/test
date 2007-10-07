<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
Delete_Id = fn_get_delete_ids
if Delete_Id<>"" then
	for each sDelete_id in split(Delete_Id,",")
		if cint(sDelete_id)<cint(50) then
		   fn_error "You are not allowed to delete system generated pages.  You may only delete pages you have created.  1 or more of the pages you selected were system pages and your request has been denied."&sDelete_id
		end if
	next

    sql_filenames = "delete from store_forms where fldpageid in ("&Delete_Id&") and store_Id="&Store_Id
	conn_store.execute sql_filenames

	server.execute "reset_design.asp"
end if

sQuestion_Path = "design/page_manager.htm"
sFlashHelp="homewysiwyg/homewysiwyg.htm"
sMediaHelp="homewysiwyg/homewysiwyg.wmv"
sZipHelp="homewysiwyg/homewysiwyg.zip"
sInstructions="Listed below are the pages used in your store. You can add additional pages or add content above or below the existing content on pre-defined pages."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Pages"
myStructure("TableWhere") = "is_link=0"
myStructure("ColumnList") = "page_id,page_name,file_name,view_order,navig_button_menu,navig_link_menu"
myStructure("HeaderList") = "page_name,file_name,view_order,navig_button_menu,navig_link_menu"
myStructure("DefaultSort") = "page_name"
myStructure("PrimaryKey") = "page_id"
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "design"
myStructure("FileName") = "page_manager.asp"
myStructure("FormAction") = "page_manager.asp"
myStructure("Title") = "Pages"
myStructure("FullTitle") = "Design > Pages"
myStructure("CommonName") = "Page"
myStructure("NewRecord") = "new_page.asp"
myStructure("Heading:page_id") = "PK"
myStructure("Heading:page_name") = "Name"
myStructure("Heading:file_name") = "File"
myStructure("Heading:view_order") = "Order"
myStructure("Heading:navig_button_menu") = "Button"
myStructure("Heading:navig_link_menu") = "Link"
myStructure("Format:page_name") = "STRING"
myStructure("Format:file_name") = "STRING"
myStructure("Format:view_order") = "INT"
myStructure("Format:navig_button_menu") = "INT"
myStructure("Format:navig_link_menu") = "INT"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->
					
<% 

createFoot thisRedirect, 0
%>
