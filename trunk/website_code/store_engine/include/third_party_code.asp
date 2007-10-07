<!--#include file="header_noview.asp"-->
<!--#include file="scart_misc.asp"-->

<%
iItem_Id = fn_get_querystring("Item_ID")

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if
show_homepage=0
bSearch=0
iItemRowPerPage=1
item_display=1
sMakeNav=0

Shopper_id=""

sql_statement="exec wsp_items_display_code "&store_Id&","&iItem_Id&",'"&Groups&"','"&Db_Full_Name&"';"

Item_L_Layout = fn_replace(Item_L_Layout,"OBJ_IMAGE_OBJ","OBJ_IMAGEL_OBJ")
Item_L_Layout = fn_replace(Item_L_Layout,"OBJ_DESCRIPTION_OBJ","OBJ_DESCRIPTIONL_OBJ")
Item_L_Layout = fn_replace(Item_L_Layout,"OBJ_NAME_OBJ","OBJ_NAME_NOLINK_OBJ<BR>")

myItem = fn_create_item(Item_L_Layout,sql_statement,bSearch,sThisUrl,iItemRowPerPage,item_display,sMakeNav)




response.write "<textarea rows=9 cols=100>"&myItem&"</textarea><BR><BR>"

%>


