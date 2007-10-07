<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

sQuestion_Path = "marketing/banners.htm"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_affiliate_banners"
myStructure("ColumnList") = "banner_id,banner_text,banner_image,view_order"
myStructure("HeaderList") = "banner_text,banner_image,view_order"
myStructure("DefaultSort") = "banner_text"
myStructure("PrimaryKey") = "banner_id"
myStructure("Level") = 7
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "marketing"
myStructure("Title") = "Marketing > Affiliates > Banners > View/Edit"
myStructure("CommonName") = "Affiliate Banner"
myStructure("NewRecord") = "affiliate_banner.asp"
myStructure("EditRecord") = "affiliate_banner.asp"
myStructure("Heading:banner_text") = "Banner Text"
myStructure("Heading:banner_image") = "Banner Image"
myStructure("Heading:view_order") = "View Order"
myStructure("Format:banner_text") = "STRING"
myStructure("Format:banner_image") = "STRING"
myStructure("Format:view_order") = "INT"

%>
<!--#include file="head_view.asp"-->

<!--#include file="list_view.asp"-->

<%
	if Request.QueryString("Delete_Id") <> "" then
	   delete_id = Request.QueryString("Delete_Id")
		sql_delete = "delete from Store_Affiliate_Banner where Banner_Id ="&delete_id&" and Store_id = "&Store_id
		conn_store.Execute sql_delete
	end if

	if  Request.Form("Delete_Id") ="SEL" then
  	if request.form("DELETE_IDS") <> "" then
  	  delete_id = request.form("DELETE_IDS")
  		sql_delete = "delete from Store_Affiliate_Banner where Banner_Id in("&delete_id&") and Store_id = "&Store_id
  		conn_store.Execute sql_delete
  	end if
  end if
set myStructure=nothing
 createFoot thisRedirect, 0%>
