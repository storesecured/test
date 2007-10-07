<!--#include file="include/header.asp"-->
<!--#include file="include/scart_misc.asp"-->

<%

fn_print_debug "search form="&request.form

sFirstPage = Switch_Name&"search_items.asp?1=1"
sPagePattern = Switch_Name&"search_items.asp?Jump_to_page=%OBJ_PAGE_OBJ%"
if request.querystring("Specials")<>"" then
	sPagePattern=sPagePattern&"&Specials=1"
end if
sUrl = replace(sPagePattern,"%OBJ_PAGE_OBJ%",fn_get_querystring("Jump_to_page"))

Show_Jump=0
bSearch =1

item_display=item_f_display
item_rows=item_f_rows
item_s_layout=item_f_layout

iItemRowPerPage=item_rows

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if


Search_Text = checkStringForQ(Request("Search_Text"))
Sub_Department_ID = fn_get_querystring("Sub_Department_ID")
Search_Dept=Request("SEARCH_DEPT")
search_MinPrice=Request("SEARCH_MINPRICE")
search_MaxPrice=Request("SEARCH_MAXPRICE")
search_SKU=checkStringForQ(Request("SEARCH_SKU"))
Jump_To_Items=checkStringForQ(fn_get_querystring("Jump_To_Items"))
specials = fn_get_querystring("Specials")
dept_full_name = checkStringForQ(request("dept_full_name"))
Sub_Department_ID = request("Sub_Department_ID")
if Search_Dept<>"" then
    Sub_Department_ID=Search_Dept
end if
Jump_to_items = fn_get_querystring("Jump_To_Items")
sString=""

if specials <> "" then
    iSpecials=1
else
    iSpecials=0
end if    

if Search_Text <> "" then
    if (len(Search_Text) > 0 and Search_Text<>"*") then
        sKeywordsArray = split(Search_Text," ")
        on error resume next
        skeyword4=sKeywordsArray(3)
        skeyword3=sKeywordsArray(2)
        skeyword2=sKeywordsArray(1)
        skeyword1=sKeywordsArray(0)
        on error goto 0
	    sString=sString&"&Search_Text="&server.URLEncode(Search_text)
    end if
end if

if dept_full_name <> "" then
    dept_full_name=dept_full_name&"%"
    sString=sString&"&dept_full_name="&server.URLEncode(dept_full_name)
end if
	
if Sub_Department_ID <> "" and Sub_Department_ID<>"-1" and isNumeric(Sub_Department_ID) then
    sString=sString&"&Sub_Department_ID="&server.URLEncode(Sub_Department_ID)
else
    Sub_Department_ID=0
end if

if search_MinPrice <> "" then
	if not isNumeric(search_MinPrice) then
		fn_redirect Switch_name&"error.asp?message_id=43"
	end if
	search_MinPrice=formatnumber(search_MinPrice,2,0,0,0)
	sString=sString&"&search_MinPrice="&server.URLEncode(search_MinPrice)
else
    search_MinPrice=-1
end if

if search_MaxPrice <> "" then
	if not isNumeric(search_MaxPrice) then
		fn_redirect Switch_name&"error.asp?message_id=43"
	end if
	search_MaxPrice=formatnumber(search_MaxPrice,2,0,0,0)
	sString=sString&"&search_MaxPrice="&server.URLEncode(search_MaxPrice)
else
    search_MaxPrice=-1
end if

if search_MinPrice <> "-1" and search_MaxPrice <> "-1" then
	if cdbl(search_MinPrice)>cdbl(search_MaxPrice) then
		fn_redirect Switch_name&"error.asp?message_id=44"
	end if
end if

if search_SKU <> "" then
    search_SKU=search_SKU&"%"
    sString=sString&"&search_SKU="&server.URLEncode(search_SKU)
end if

Jump_To_Items=checkStringForQ(fn_get_querystring("Jump_To_Items"))
if Jump_to_items <> "" then
	Jump_To_Items=Jump_To_Items&"%"
	sString=sString&"&Jump_To_Items="&server.URLEncode(Jump_To_Items)
end if
	

show_homepage=0
q_Item_Page_Name=""
sMakeNav=1

if Show_TopNav then
	sBreadcrumbs = "<table border='0' cellspacing='0' width='100%'><tr><td>"
	if Top_Depts="" then
		Top_Depts="Top"
	end if
	sBreadcrumbs = sBreadcrumbs & ("<b><a href='"&fn_dept_url("","")&"' class=link>"&Top_Depts&"</a></b>")
	sBreadcrumbs = sBreadcrumbs & ("<b> > <a href='"&fn_page_url("search.asp",0)&"' class=link>Search</a></b>")
	sBreadcrumbs = sBreadcrumbs & ("<b> > Search Results</b>")
	sBreadcrumbs = sBreadcrumbs & ("<hr></td></tr></table>")

	response.write sBreadcrumbs
end if

if sString<>"" then
    if instr(sUrl,"?")>0 then
        sUrl = sUrl&sString
    else
        sUrl = sUrl&"?1=1"&sString
    end if
    
    sDetailUrl =  sUrl
    sThisUrl = sUrl 

    sFirstPage = sFirstPage&sString
    sPagePattern = sPagePattern&sString

end if

sql_statement="exec wsp_items_search "&store_id&","&sHide_Out_stock&",'%"&skeyword1&"%','%"&skeyword2&"%','%"&skeyword3&"%','%"&skeyword4&"%',"&iSpecials&",'"&dept_full_name&"',"&search_MinPrice&","&search_MaxPrice&",'"&search_SKU&"','"&Jump_To_Items&"',"&sub_department_id&",'"&Groups&"';"
fn_print_debug sql_statement

item_f_layout = fn_replace(item_f_layout,"OBJ_IMAGE_OBJ","OBJ_IMAGES_OBJ")
item_f_layout = fn_replace(item_f_layout,"OBJ_DESCRIPTION_OBJ","OBJ_DESCRIPTIONS_OBJ")
item_f_layout = fn_replace(item_f_layout,"OBJ_EXT_FIELD2_OBJ","")
item_f_layout = fn_replace(item_f_layout,"OBJ_EXT_FIELD3_OBJ","")
item_f_layout = fn_replace(item_f_layout,"OBJ_EXT_FIELD4_OBJ","")
item_f_layout = fn_replace(item_f_layout,"OBJ_EXT_FIELD5_OBJ","")
sItemText = fn_create_item(item_f_layout,sql_statement,bSearch,sThisUrl,iItemRowPerPage,item_display,sMakeNav)
response.Write sItemText

%>

<!--#include file="include/footer.asp"-->