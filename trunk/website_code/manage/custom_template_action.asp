<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->


<%
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%
else
    template_name = checkStringForQ(Request.Form("template_name"))
    template_description = checkStringForQ(Request.Form("template_description"))
	
	if request.form("op") <> "edit" then	
		sql_select="exec wsp_design_create_blank "&Store_Id&",'"&template_name&"','"&template_description&"'"
	    conn_store.execute(sql_select)
	    response.write sql_select
	    
	    sql="select max(template_id) as new_template_id from store_design_template where Store_Id="&Store_Id
		rs_store.open sql,conn_store,1,1
		if not rs_store.EOF then
			template_id=rs_store("new_template_id")
		end if
		rs_store.close
		response.write sql
	 else
		template_id = request.form("template_id")
        sql_add = "update store_design_template set template_name='"&template_name&"',template_description='"&template_description&"' where Store_Id="&Store_Id & " and template_id= " & template_id&" and template_active=0"
        conn_store.Execute(sql_add)
    end if
    Response.redirect "layout_design.asp?Id="&Template_Id
    
end if
%>

