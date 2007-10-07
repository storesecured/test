<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "marketing/banners.htm"
sTextHelp="banners/banners.doc"

sInstructions="Insert the keyword %OBJ_BANNER_OBJ% where you want the banner to appear."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_banners"
myStructure("ColumnList") = "bann_id,banner_name,enabled"
myStructure("HeaderList") = "banner_name,enabled,clicks,impressions"
myStructure("DefaultSort") = "banner_name"
myStructure("PrimaryKey") = "bann_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "banners.asp"
myStructure("FormAction") = "banners.asp"
myStructure("Title") = "Banners"
myStructure("FullTitle") = "Marketing > Banners"
myStructure("CommonName") = "Banner"
myStructure("NewRecord") = "banners_add.asp"
myStructure("Heading:bann_id") = "PK"
myStructure("Heading:enabled") = "Enabled"
myStructure("Heading:banner_name") = "Name"
myStructure("Heading:clicks") = "Clicks"
myStructure("Heading:impressions") = "Impressions"
myStructure("Format:banner_name") = "STRING"
myStructure("Format:clicks") = "TEXT"
myStructure("Format:impressions") = "TEXT"
myStructure("Format:enabled") = "LOOKUP"
myStructure("Link:clicks") = "banners_rc.asp?Id=PK"
myStructure("Link:impressions") = "banners_ri.asp?Id=PK"
myStructure("Lookup:enabled") = "False:No^True:Yes"
'myStructure("Sql:clicks") = "select count(Banner_Click_Id) as clicks from store_banners_click_through where store_id="&Store_Id&" and banner_id=PK"
'myStructure("Sql:impressions") = "select count(Banner_Impression_Id) as impressions from store_banners_impressions where store_id="&Store_Id&" and banner_id=PK"


%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
	
Delete_Id = fn_get_delete_ids

if Delete_Id<>"" then
    sql_del = "delete from store_banners_click_through where Banner_ID in ("&Delete_Id&") and store_id="&store_id
    conn_store.execute sql_del
    sql_del = "delete from store_banners_impressions where Banner_ID in ("&Delete_Id&") and store_id="&store_id
    conn_store.execute sql_del
end if
createFoot thisRedirect, 0
%>


