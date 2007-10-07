<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->


<%
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><%

else
  if request.querystring("Delete") <> "" then
	  if isNumeric(request.querystring("Delete")) then
	  on error goto 0
	  sql_delete = "delete from Store_Pages where Store_Id = " &Store_Id & " and Page_Id = " & request.querystring("Delete")
	  conn_store.Execute sql_delete
	  sql_delete = "delete from Store_Page_Content where Store_Id = " &Store_Id & " and Page_Id = " & request.querystring("Delete")
	  conn_store.Execute sql_delete
	  end if
  else

	 Custom_Field_Name = checkStringForQ(Request.Form("Custom_Field_Name"))
	 if Custom_Field_Name = "" then
		 response.redirect "admin_error.asp?message_id=21"
	 end if

	 if request.form("op") <> "edit" then

'		sql="select max(CustomField_ID)+1 as new_CustomField_Id from store_custom_fields where Store_Id="&Store_Id
'		rs_store.open sql,conn_store,1,1
'			CustomField_ID=rs_store("new_CustomField_Id")
'		rs_store.close

	 else
		CustomField_ID = request.form("CustomField_ID")

	 end if

 
	 Custom_Field_Type = Request.Form("Custom_Field_Type")

	if Custom_Field_Type = 1 then
		Custom_Field_Values = Request.Form("Custom_Text_Value")
		if not isNumeric(Custom_Field_Values) then
			fn_error "Please enter the size of the text field."
		end if
	elseif Custom_Field_Type = 2 then
		Custom_Field_Values = Request.Form("Custom_Text_Area_Value")
		if not isNumeric(Custom_Field_Values) then
			fn_error "Please enter the number of columns for the textarea."
		end if
	elseif  Custom_Field_Type = 3 then
		Custom_Field_Values = checkStringForQ(Request.Form("Custom_Combo_Value"))
	end if
		
	
	'if Custom_Field_Type = 3 then
  	'	 Custom_Field_Values = checkStringForQ(Request.Form("Custom_Field_Values"))
	'else
	'	 Custom_Field_Values = Request.Form("Custom_Field_Values")
	'end if

	 if Request.Form("Field_Required") <> "" then
		 Field_Required = 1
	 else
		 Field_Required = 0
	 end if

	 View_Order = Request.Form("View_Order")

	 if request.form("op") <> "edit" then
		'ADD FIELD TO STORE_CUSTOM_FIELDS TABLE
		sql_add = "insert into Store_Custom_Fields (Store_id,Custom_Field_Name,Custom_Field_Type, Custom_Field_Values,Field_Required,View_Order) values ("&Store_id&",'"&Custom_Field_Name&"','"&Custom_Field_Type&"', '"&Custom_Field_Values&"',"&Field_Required&","&View_Order&")"

	 
	 else
		 if View_Order = "" then
			 sql_add = "update Store_Custom_Fields set Custom_Field_Name='"&Custom_Field_Name&"',Custom_Field_Type='"&Custom_Field_Type&"',Custom_Field_Values='"&Custom_Field_Values&"',Field_Required="&Field_Required&" where Store_Id="&Store_Id & " and CustomField_ID= " & CustomField_ID
		 else
			 sql_add = "update Store_Custom_Fields set Custom_Field_Name='"&Custom_Field_Name&"',Custom_Field_Type='"&Custom_Field_Type&"',Custom_Field_Values='"&Custom_Field_Values&"',Field_Required="&Field_Required&",View_Order="&View_Order&" where Store_Id="&Store_Id & " and CustomField_ID= " & CustomField_ID

		 end if
	 end if

	 conn_store.Execute sql_add

  end if

  
	returnTo = Request.ServerVariables("HTTP_REFERER")
          if returnTo = "" then
             response.redirect "fields.asp"
          end if
 response.redirect returnTo

end if
%>

