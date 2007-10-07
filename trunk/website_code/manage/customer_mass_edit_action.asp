<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
'USER CLICKED ON "MASS EDIT"
'UPDATE ALL SELECTED ITEMS ACCORDING TO THE MODIFICATIONS MADE BY USER
if request.form("Form_Name")= "Mass_Edit_Update" then

	If Form_Error_Handler(Request.Form) <> "" then
	   Error_Log = Form_Error_Handler(Request.Form)
	   if Error_Log <> "" then
		%> <!--#include virtual="common/Error_Template.asp"--><%
		response.end
		end if
	end if

	'GET THE VARIABLES

	if request.form("spam")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " spam="&request.form("spam")
	end if

	if request.form("tax_exempt")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " tax_exempt="&request.form("tax_exempt")
	end if

	if request.form("Protected_Page_Access")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Protected_Page_Access="&request.form("Protected_Page_Access")
	end if

	Budget_orig = request.form("budget")

	if request.form("budget")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Budget_orig="&request.form("budget") 
	end if


	if request.form("budget_left")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " budget_left="&request.form("budget_left") 
	end if



if trim(updates_str) <> "" then
	sql_str =  "update store_customers set " & updates_str  & " where Store_ID = " & store_id &" and  Ccid in (" & request.form("delete_ids") & ")"
	conn_store.Execute sql_str

else
	response.redirect "error.asp?message_id=94"
end if

	
End if

response.redirect "my_customer_base.asp"

%>
