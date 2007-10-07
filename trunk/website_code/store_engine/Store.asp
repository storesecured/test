<!--#include file="include/header.asp"-->
<!--#include file="include/scart_misc.asp"-->

<%

sFirstPage = Switch_Name
sPagePattern = fn_dept_url("homepage","%OBJ_PAGE_OBJ%")

sUrl = fn_dept_url("homepage",fn_get_querystring("Jump_to_page"))
sThisUrl = sUrl

if item_rows = "" then
	item_rows = 10
end if

if not isNumeric(item_f_rows) or item_f_rows>50 then
   item_f_rows=item_rows
end if

bSearch =0

item_display=item_f_display
item_rows=1000
item_s_layout=item_f_layout

iItemRowPerPage = item_rows

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if

Jump_To_Items=checkStringForQ(fn_get_querystring("Jump_To_Items"))

dept_full_name=""
show_homepage=1
q_Item_Page_Name=""
sMakeNav=1

sql_statement="exec wsp_items_display_homepage "&store_Id&","&sHide_Out_stock&",'"&groups&"';"

item_s_layout = fn_replace(item_s_layout,"OBJ_IMAGE_OBJ","OBJ_IMAGES_OBJ")
item_s_layout = fn_replace(item_s_layout,"OBJ_DESCRIPTION_OBJ","OBJ_DESCRIPTIONS_OBJ")

sItemText = fn_create_item(item_s_layout,sql_statement,1,sThisUrl,iItemRowPerPage,item_display,1)
response.Write sItemText

%>

<!--#include file="include/footer.asp"-->
