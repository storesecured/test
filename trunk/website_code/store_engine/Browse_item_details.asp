<!--#include file="include/header.asp"-->
<!--#include file="include/scart_misc.asp"-->
<%

sUrl =  fn_dept_url(Db_Full_Name,"")
sDetailUrl =  fn_item_url(Db_Full_Name,q_item_page_name)
sThisUrl = sDetailUrl 

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if

show_homepage=0
if Detail_NextPrev then
	item_display = item_display+1
	item_rows = item_rows
	if item_rows = "" then
		item_rows = 10
	end if
    
    prev_item_page_name=""
    next_item_page_name=""
	sql_select = "exec wsp_items_pagename "&store_Id&","&sHide_Out_stock&","&iDepartment_Id&";"
	fn_print_debug sql_select
	set browsefields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,browsedata,browsefields,noRecordsbrowse)
	if noRecordsbrowse = 0 then
	    FOR rowcounterbrowse = 0 TO browsefields("rowcount")
            this_item_page_name = browsedata(browsefields("item_page_name"),rowcounterbrowse)
            if this_item_page_name=q_Item_Page_Name then
                if browsefields("rowcount")>=(rowcounterbrowse + 1) then
                    next_item_page_name = browsedata(browsefields("item_page_name"),rowcounterbrowse + 1)
                end if
                exit for
            end if
            prev_item_page_name = this_item_page_name
	    next
	end if

    sUpButton = fn_create_action_button ("Button_image_Up", "Next_Jump_Click_Items", "Up")
	sUpButton = "<form action='"&sUrl&"' method='post'>"&_
		"<input type='hidden' name='Current_Start_Row_Items_Items' value='"&(rowcounterbrowse-(item_rows*item_display))&"'>"&_
		"<input type='hidden' name='Current_End_Row_Items_Items' value='"&rowcounterbrowse&"'>"&_
		"<input type='hidden' name='Next_jump_Items' size='3' value='"&(item_rows*item_display)&"'>"&_
		sUpButton&"</form>"

	if prev_item_page_name <> "" then
		sPrevButton = fn_create_action_button ("Button_image_Prev", "Back_Jump_Click_Items", "Prev")
		sRecords = sRecords & "<td align=right width='45%'><form action='"&fn_item_url(Db_Full_Name,prev_item_page_name)&"' method='post'>"&sPrevButton&"</form></td>"
	else
		sRecords = sRecords & "<td align=right width='45%'>&nbsp;</td>"
	end if
	sRecords = sRecords & "<td align=center width='10%'>"&sUpButton&"</td>"
	if next_item_page_name <> "" then
		sNextButton = fn_create_action_button ("Button_image_Next", "Next_Jump_Click_Items", "Next")
		sRecords = sRecords & "<td align=left width='45%'><form action='"&fn_item_url(Db_Full_Name,next_item_page_name)&"' method='post'>"&sNextButton&"</form></td>"
	else
		sRecords = sRecords & "<td align=left width='45%'>&nbsp;</td>"
	end if
end if

response.write fn_show_breadcrumbs (Show_TopNav,Db_Full_Name,q_item_page_name)

if sRecords <> "" then
		response.write "<table border=0 cellspacing=0 cellpadding=0 width='100%'><TR>"&sRecords&"</tr></table>"
end if

show_homepage=0
bSearch=0
iItemRowPerPage=1
item_display=1
sMakeNav=0

sql_statement="exec wsp_items_display_detail "&store_Id&","&sHide_Out_stock&","&iItem_Id&",'"&Groups&"','"&Db_Full_Name&"';"

Item_L_Layout = fn_replace(Item_L_Layout,"OBJ_IMAGE_OBJ","OBJ_IMAGEL_OBJ")
Item_L_Layout = fn_replace(Item_L_Layout,"OBJ_DESCRIPTION_OBJ","OBJ_DESCRIPTIONL_OBJ")
Item_L_Layout = fn_replace(Item_L_Layout,"OBJ_NAME_OBJ","OBJ_NAME_NOLINK_OBJ<BR>")

myItem = fn_create_item(Item_L_Layout,sql_statement,bSearch,sThisUrl,iItemRowPerPage,item_display,sMakeNav)
response.Write myItem

if sRecords <> "" then
		response.write "<table border=0 cellspacing=0 cellpadding=0 width='100%'><TR>"&sRecords&"</tr></table>"
end if
%>

<!--#include file="include/footer.asp"-->
