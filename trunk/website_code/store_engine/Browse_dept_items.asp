<!--#include file="include/header.asp"-->
<!--#include file="include/scart_misc.asp"-->

<% 

server.scripttimeout=10

fn_print_debug "q_Dept_Page_Name in browse="&q_Dept_Page_Name

sFirstPage = fn_dept_url(Db_Full_Name,"")
sPagePattern = fn_dept_url(Db_Full_Name,"%OBJ_PAGE_OBJ%")

sThisUrl = fn_dept_url(Db_Full_Name,fn_get_querystring("Jump_to_page"))

dept_full_name=fn_transform_deptstring(q_Dept_Page_Name)

if dept_rows = "" then
	dept_rows = 10
end if
if item_rows = "" then
	item_rows = 10
end if
iDeptRowPerPage = dept_rows

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if

searchOn=0
sql_select_items = "exec wsp_dept_display_detail "&Store_Id&","&iDepartment_Id&";"
fn_print_debug sql_select_items
session("sql")=sql_select_items
rs_store.open sql_select_items,conn_store,1,1
if rs_store.eof then
    'dept doesnt exist
    fn_print_debug sql_select_items
    fn_print_debug "dept doesnt exist, redirecting"
    fn_redirect fn_dept_url("","")
else
    this_Department_ID=rs_store("Department_ID")	
    this_Department_Name=rs_store("Department_Name")	
    this_Department_HTML=rs_store("Department_HTML")	
    this_Department_HTML_Bottom=rs_store("Department_HTML_Bottom")	
    this_Last_Level=rs_store("Last_Level")
    this_show_name=rs_store("show_name")
end if
rs_store.close

sDeptHead = ""
sDeptHead = sDeptHead& fn_show_breadcrumbs (Show_TopNav,Db_Full_Name,"")

if not isNull(this_Department_HTML) and this_Department_HTML<>"" then
    sDeptHead = sDeptHead&("<!-- department html top start--><table width='100%' border=0 cellspacing=0 cellpadding=0><tr><td>"&_
	    fn_putCustomHTML(this_Department_HTML)&_
	    "</td></tr></table><!-- department html top end -->")
end if

Jump_To=checkStringForQ(fn_get_querystring("Jump_To"))

sql_statement = "exec wsp_dept_display_browse "&Store_Id&","&this_Department_Id&",'"&Jump_To&"';"
sDeptStr = fn_create_dept (Department_Layout,q_Dept_Page_Name,sql_statement,Hide_Empty_Depts,sUrl,iDeptRowPerPage,dept_display,number_of_recordset,Dept_Start_Row,Dept_End_Row)

if iDepartment_Id<>0 then
	sItemString = fn_browse_items()
end if

if Subdept_location then
	sDeptText = sDeptHead&sDeptStr&sItemString
else
    sDeptText = sDeptHead&sItemString&sDeptStr
end if

if not isNull(this_Department_HTML_Bottom) and this_Department_HTML_Bottom<>"" then
	sDeptText = sDeptText&(fn_putCustomHTML(this_Department_HTML_Bottom))
end if

response.Write sDeptText

set deptfields1 = Nothing

function fn_browse_items ()
    if not isNumeric(item_f_rows) or item_f_rows>50 then
       item_f_rows=item_rows
    end if

    bSearch =0
    iItemRowPerPage = item_rows

    Jump_To_Items=checkStringForQ(fn_get_querystring("Jump_To_Items"))

    show_homepage=0
    q_Item_Page_Name=""
    sMakeNav=1

    sql_statement="exec wsp_items_display_browse "&store_Id&","&sHide_Out_stock&","&iDepartment_Id&",'"&checkstringforQ(Db_Full_Name)&"','"&Jump_To_Items&"','"&groups&"';"

    item_s_layout = fn_replace(item_s_layout,"OBJ_IMAGE_OBJ","OBJ_IMAGES_OBJ")
    item_s_layout = fn_replace(item_s_layout,"OBJ_DESCRIPTION_OBJ","OBJ_DESCRIPTIONS_OBJ")
    item_s_layout = fn_replace(item_s_layout,"OBJ_ACCESSORY_OBJ","")
    myItems = fn_create_item(item_s_layout,sql_statement,bSearch,sThisUrl,iItemRowPerPage,item_display,sMakeNav)
    
    fn_browse_items=myItems
end function

%>
<!--#include file="include/footer.asp"-->
