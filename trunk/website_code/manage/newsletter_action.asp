<!--#include file="Global_Settings.asp"-->
<%
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><% 
	
else
	
	'CREATE / EDIT PROMOTION
	if Request.Form("Form_Name") = "Create_Newsletter" then
		'RETRIEVE FORM DATA
		Email_Address = checkStringForQ(request.form("Email_Address"))
		First_Name = checkStringForQ(request.form("First_Name"))
		Last_Name = checkStringForQ(request.form("Last_Name"))

          Newsletter_Id = request.form("Newsletter_Id")

		if Request.Form("op")<>"edit" then
			'IF ADD, THEN PREPARE AND RUN AN INSERT
			sql_insert = "insert into store_newsletter (Store_id,Email_Address,First_Name,Last_Name) values ("&Store_id&",'"&Email_Address&"','"&First_Name&"','"&Last_Name&"')"
			session("sql") = sql_insert
         conn_store.Execute sql_insert
			response.redirect "Newsletter_manager.asp"

		else
			'IF EDIT, THEN PREPARE AND RUN AN UPDATE
			Newsletter_Id=Request.Form ("Newsletter_Id")
			sqlUpdate="update store_newsletter set Email_Address='" & Email_Address & "', " & _
			          "First_Name='"&First_Name&"',Last_Name='"&Last_Name&"' "&_
					"where store_id=" & store_id & " and newsletter_id=" & newsletter_id
			conn_store.Execute sqlUpdate

			response.redirect "Newsletter_manager.asp"

		end if
	end if

	'DELETE PROMOTION
	if Request.QueryString("Delete_Newsletter_Id") <> "" then 
		Newsletter_Id = Request.QueryString("Delete_Newsletter_Id")
		if not isNumeric(Newsletter_Id) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		sql_delete = "delete from store_newsletter where Newsletter_Id = "&Newsletter_Id
		conn_store.Execute sql_delete
		response.redirect "Newsletter_manager.asp"
		
	end if

End If

%> 
