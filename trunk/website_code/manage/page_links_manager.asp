<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
sQuestion_Path = "design/page_links_manager.htm"
sInstructions="Listed below are the External Links used in your store. You can add additional pages or add content above or below the existing content on pre-defined pages."
		
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Pages"
myStructure("TableWhere") = "is_link=1"
myStructure("ColumnList") = "page_id,page_name,view_order,navig_button_menu,navig_link_menu"
myStructure("HeaderList") = "page_name,view_order,navig_button_menu,navig_link_menu"
myStructure("DefaultSort") = "page_name"
myStructure("PrimaryKey") = "page_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "design"
myStructure("FileName") = "page_links_manager.asp"
myStructure("FormAction") = "page_links_manager.asp"
myStructure("Title") = "Links"
myStructure("FullTitle") = "Design > Links"
myStructure("CommonName") = "Link"
myStructure("NewRecord") = "new_page_link.asp"
myStructure("Heading:page_id") = "PK"
myStructure("Heading:page_name") = "Name"
myStructure("Heading:view_order") = "Order"
myStructure("Heading:navig_button") = "Button"
myStructure("Heading:navig_link") = "Link"
myStructure("Format:page_name") = "STRING"
myStructure("Format:view_order") = "INT"
myStructure("Format:navig_button_menu") = "INT"
myStructure("Format:navig_link_menu") = "INT"

%>
<!--#include file="head_view.asp"-->

<!--#include file="list_view.asp"-->
		
<%
createFoot thisRedirect, 0
%>
