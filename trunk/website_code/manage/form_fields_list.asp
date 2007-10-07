<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="createForm.asp"-->
<% 

fldPBID = Request.QueryString("fldPBID")
if len(fldPBID)  = 0 then
	response.Redirect("admin_error.asp?message_id=105")
end if
		if isNumeric(fldPBID) then
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		
		
		set myStructure=server.createobject("scripting.dictionary")
		myStructure("TableName") = "Store_Form_Fields"
		myStructure("TableWhere") = "fldpbid="&request.querystring("fldPBID")
		myStructure("ColumnList") = "fldid,fldname,fldvieworder,fldfieldtype"
		myStructure("DefaultSort") = "fldvieworder"
		myStructure("PrimaryKey") = "fldid"
		myStructure("Level") = 1
		myStructure("EditAllowed") = 1
		myStructure("AddAllowed") = 1
		myStructure("DeleteAllowed") = 1
'		myStructure("BackTo") = "fields_page.asp?Id="&request.querystring("Item_Id")
		myStructure("BackTo") = "fields_page.asp"
		myStructure("BackToName") = "Form"
		myStructure("Menu") = "design"
		myStructure("FileName") = "form_fields_list.asp?fldPBID="&request.querystring("fldPBID")
		myStructure("FormAction") = "form_fields_list.asp?fldPBID="&request.querystring("fldPBID")
		myStructure("Title") = "Forms Fields"
		myStructure("FullTitle") = "Design > <a href=fields_page.asp class=white>Forms</a> > Fields"
		myStructure("CommonName") = "Form Field"
		myStructure("NewRecord") = "form_fields_add.asp?id="&request.querystring("fldPBID")
		myStructure("EditRecord") ="form_fields_add.asp"&request.querystring("Item_Id")
		myStructure("Heading:fldid") = "PK"
		myStructure("Heading:fldname") = "Field Name"
		myStructure("Heading:fldvieworder") = "View Order"
		myStructure("Heading:fldfieldtype") = "Field Type"

		myStructure("Format:fldname") = "STRING"
		myStructure("Format:fldvieworder") = "INT"
		myStructure("Format:fldfieldtype") = "LOOKUP"
		myStructure("Lookup:fldfieldtype") = "1:Text Field^2:Text Area^3:Combo Box^4:Check Box"
	
		myStructure("Length:fldname") = "25"
'		myStructure("Length:fldvieworder") = "4"

'		myStructure("Format:customer_group") = "SQL"
'		myStructure("Format:group_price") = "CURR"

'		myStructure("Length:customer_group") = "0"
'		myStructure("Sql:customer_group") = "select group_name from Store_Customers_Groups where group_id in (select group_id from Store_Items_Price_Group where item_id =" & request.querystring("Item_Id")
	'	myStructure("Sql:customer_group") = "select group_name as customer_group from Store_Customers_Groups where group_id = THISFIELD and store_id="&Store_Id

	'	myStructure("Length:group_price") = "0"

		
	
		
		
		


		
		
%>


<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->
<%

Delete_Id = fn_get_delete_ids
if Delete_Id<>"" then
    call generateForm(fldPBID)
	response.redirect "form_fields_list.asp?fldPBid=" & fldPBID
end if

createFoot thisRedirect, 0%>
