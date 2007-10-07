<!--#include file="Global_Settings.asp"-->
<%
	page_id = request.querystring("page_id")
	file_name = replace(request.querystring("file_name"),".asp","_copy.asp")

	sql_copy = "exec wsp_page_copy " & Store_Id &","& page_Id&",'"&file_name&"';"
	conn_store.Execute sql_copy
	server.execute "reset_design.asp"
	
    response.redirect "page_manager.asp"
%>

