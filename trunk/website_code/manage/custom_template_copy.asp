<!--#include file="Global_Settings.asp"-->

<%
Template_Id = request.querystring("Id")
if Template_Id = "" then
	fn_error "Select a template to copy"
else	
	sql_copy = "exec wsp_design_copy "&Store_Id&","&Store_Id&","&Template_Id
	conn_store.Execute sql_copy
	
	sql_select = "select max(template_id) as template_id_copy from store_design_template where store_id="&Store_Id
	rs_Store.open sql_select,conn_store,1,1
	if not rs_store.eof then
        Template_Id_Copy=rs_store("Template_Id_Copy")
    end if
    rs_store.close

	response.redirect "layout_design.asp?op=edit&Id="&Template_Id_Copy
end if

%>
