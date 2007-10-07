
<!--#include file="store_settings.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<!--#include file="create_objects.asp"-->
<%

if CurrentFileName = "browse_dept_items.asp" then
    q_Dept_Page_Name = fn_get_querystring("dept_page_name")
    if q_Dept_Page_Name="" then
        categ_id=fn_get_querystring("categ_id")
        if categ_id<>"" then
            q_Dept_Page_Name = fn_old_dept_url(categ_id)
        end if
    end if
    Full_Name=fn_transform_deptstring(q_Dept_Page_Name)
    q_item_page_name=""
    q_page_name=CurrentFileName
elseif CurrentFileName = "browse_item_details.asp" then
    q_item_page_name = fn_get_querystring("item_page_name")
    q_Dept_Page_Name = fn_get_querystring("dept_page_name")
    if q_item_page_name="" then
        categ_id = fn_get_querystring("categ_id")
        item_id=fn_get_querystring("item_id")
        if item_id<>"" then
            q_item_page_name = fn_old_item_url (categ_id,Item_ID)
        end if
    end if
    full_name = fn_transform_deptstring(q_Dept_Page_Name)
    q_page_name=CurrentFileName
elseif CurrentFileName = "custom.asp" then
    q_page_name = fn_get_querystring("page_name")
    if q_page_name="" then
        page_id=fn_get_querystring("page_id")
        if page_id<>"" then
            q_page_name = fn_old_page_url (page_id)
        end if
    end if
else
	q_page_name = CurrentFileName
end if

q_item_page_name=checkstringforQ(q_item_page_name)
q_page_name=checkstringforQ(q_page_name)
Full_Name=checkstringforQ(Full_Name)
iUse_Template=-1

sql_select = "exec wsp_page_info "&Store_Id&",'"&q_page_name&"','"&Full_Name&"','"&q_item_page_name&"';"
fn_print_debug sql_select
rs_store.open sql_select,conn_store,1,1
if not rs_store.eof then
	Db_Full_Name = rs_store("Full_Name")
	on error resume next
	iDepartment_Id = rs_store("Department_Id")
	iItem_Id = rs_store("Item_Id")
	sMeta_Keywords = rs_store("Meta_Keywords")
	sMeta_Description = rs_store("Meta_Description")
	sMeta_Title = rs_store("Meta_Title")
	sCustomer_Group=rs_store("Customer_Group")
	iUse_Template=cint(rs_store("Use_Template"))
	sProtect_page = rs_store("Protect_Page")
	sPage_form_content=rs_store("page_form_content")
	sPage_top=rs_store("page_top")
	sPage_bottom=rs_store("page_bottom")
	sPage_Head=rs_store("page_head")

	if sMeta_Title="" then
		'default page title to the name of the item or dept
		sMeta_Title=fn_stripHTML(rs_store("myName"))
	end if
	if sMeta_Description="" then
		'default page title to the name of the item or dept
		sMeta_Description=fn_stripHTML(rs_store("myDescription"))
	end if
	rs_store.close

	if sProtect_page then
		fn_page_protected
	end if
	if sCustomer_Group<>"0" and isNumeric(sCustomer_Group) and sCustomer_Group<>"" then
		fn_page_group
	end if
	if q_Dept_Page_Name<>"" and instr(q_Dept_Page_Name,"/")>0 then
		'only need to check on departments that arent top level
		iProtected=fn_check_protected(iDepartment_Id,Db_Full_Name)
	end if
elseif CurrentFileName = "browse_item_details.asp" then
	'item didnt exist in this dept try to find it in another dept
	rs_store.Close
	fn_item_missing full_name, q_item_page_name
elseif CurrentFileName = "browse_dept_items.asp" then
    rs_store.Close
    if q_Dept_Page_Name<>"Un Assigned" and q_Dept_Page_Name<>"" then
	    fn_print_debug "in dept not equal unassigned"
        'item or dept didnt exist
        fn_dept_missing full_name
    else
        q_Dept_Page_Name="Un Assigned"
    end if
else
    rs_store.Close
    fn_redirect Switch_name&"404.asp"
end if

if iDepartment_Id="" then
	iDepartment_Id=0
end if
if iItem_Id="" then
	iItem_Id=0
end if

if sProtect_page and Protected_page_access=0 and CurrentFileName <> "check_out.asp" then
    fn_print_debug "page is protected"
	fn_redirect Switch_Name&"check_out.asp?Protected=Yes&ReturnTo="&server.urlencode(CurrentFilename&"?"&request.querystring)
end if

sCustomer_Group=Only_Customer_Group
if sCustomer_Group<>"0" and isNumeric(sCustomer_Group) and sCustomer_Group<>"" and CurrentFileName <> "check_out.asp" and not(Is_In_Collection(Groups,sCustomer_Group,",")) then
	fn_print_debug "sCustomer_Group="&sCustomer_Group
	fn_print_debug "Groups="&Groups
	fn_redirect Switch_Name&"check_out.asp?Protected=Yes&ReturnTo="&server.urlencode(CurrentFilename&"?"&request.querystring)
end if

on error goto 0

if fn_get_querystring("Preview")="" then
	sPreview=0
else
	sPreview=1
end if
sql_select = "exec wsp_template_display "&Store_Id&","&sPreview&","&iUse_Template&";"
fn_print_debug sql_select
rs_store.open sql_select,conn_store,1,1
if not rs_store.eof then
    sTemplate_html=rs_store("template_html")
    Button_image_Continue = rs_Store("Button_image_Continue")
	Button_image_Update = rs_Store("Button_image_Update")
	Button_image_Up = rs_Store("Button_image_Up")
	Button_image_Cancel = rs_Store("Button_image_Cancel")
	Button_image_SaveCart = rs_Store("Button_image_SaveCart")
	Button_image_RetrieveCart = rs_Store("Button_image_RetrieveCart")
	Button_image_Search = rs_Store("Button_image_Search")
	Button_image_Delete = rs_Store("Button_image_Delete")
	Button_image_View = rs_Store("Button_image_View")
	Button_image_Login = rs_Store("Button_image_Login")
	Button_image_ProcessOrder = rs_Store("Button_image_ProcessOrder")
	Button_image_No = rs_Store("Button_image_No")
	Button_image_Yes = rs_Store("Button_image_Yes")
	Button_image_UpdateCart = rs_Store("Button_image_UpdateCart")
	Button_image_Checkout = rs_Store("Button_image_Checkout")
	Button_image_Order = rs_Store("Button_image_Order")
	Button_image_CancelOrder = rs_Store("Button_image_CancelOrder")
	Button_image_Reset = rs_Store("Button_image_Reset")
    Button_image_Next = rs_Store("Button_image_Next")
    Button_image_Prev = rs_Store("Button_image_Prev")
    Button_image_ContinueShopping = rs_Store("Button_image_ContinueShopping")
    Template_Buttons = rs_Store("Buttons")
    Template_Links = rs_Store("Links")
end if
rs_store.close

fn_get_template=fn_putCustomHTML(sTemplate_html)

cPos = InStr(fn_get_template, "%OBJ_CENTER_CONTENT_OBJ%")
if cPos > 0 then
	fPos = Len(fn_get_template)
	Store_Header = Left(fn_get_template,cPos-1)
	leftPos = fPos - cPos
	Store_Footer = Right(fn_get_template,leftPos-23)
else
    Store_Header = fn_get_template
    Store_Footer = ""
end if


response.Write Store_Header

%>








