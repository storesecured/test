<!--#include file="Global_Settings.asp"-->
<%
	'CHECK THE MODE
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

if request.form="" then
	response.redirect "page_links_manager.asp"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%> <!--#include file="Include/Error_Template.asp"--><% 
else
	
	page_name = checkStringForQ(request.form("page_name"))
	file_name =  checkStringForQ(request.form("file_name"))
	page_description = checkStringForQ(request.form("page_description"))
	view_order = request.form("view_order")

	 Navig_Link_Menu = Request.Form("Navig_Link_Menu")
     Navig_Button_Menu = Request.Form("Navig_Button_Menu")
	 page_id = request.form("page_id")

	if request.form("op") = "edit" and page_id<>"" then
		Sql_update  = "update store_pages set page_name = '" & page_name & "' ,  file_name = '" & file_name & "' , page_description = '" & page_description & "' , navig_button_menu = " & navig_button_menu & " , navig_link_menu = " & navig_link_menu &   " , view_order = " & view_order &" where store_id = " & store_id & " and page_id = " & page_id
	   session("sql")=Sql_update
		conn_store.Execute sql_update
	else
		sql_insert = "exec wsp_page_link " & store_id & ", '" & page_name &"', '" & file_name &"', '" & page_description & "', " & view_order  & " , " & navig_button_menu & " , " & Navig_Link_Menu
		session("sql")=sql_insert
		conn_store.Execute sql_insert
	end if
	if cint(request.form("Navig_Button_Menu_Old"))<>cint(Navig_Button_Menu) or cint(request.form("Navig_Link_Menu_Old")) <> cint(Navig_Link_Menu) then
        'only recreate the design if something needs to be added to a menu
        server.execute "reset_design.asp"
    end if


	response.redirect "page_links_manager.asp"
end if	
%>
