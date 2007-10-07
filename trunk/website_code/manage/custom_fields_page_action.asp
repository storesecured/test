<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include file="createForm.asp"-->
<%
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)%>
	<!--#include file="Include/Error_Template.asp"-->
<%else
 if request.querystring("Delete") <> "" then
	if isNumeric(request.querystring("Delete")) then
		on error goto 0
		sql_delete = "delete from Store_Pages where Store_Id = " &Store_Id & " and Page_Id = " & request.querystring("Delete")
		'''ust= conn_store.Execute sql_delete
		sql_delete = "delete from Store_Page_Content where Store_Id = " &Store_Id & " and Page_Id = " & request.querystring("Delete")
		'''ust= conn_store.Execute sql_delete
	end if
 else
	
	pageselect=Request.Form("pageselect")
	Emailto =Request.Form("Emailto")
	Subject =Request.Form("Subject")
	varid= Request.Form("CustomField_ID")

	str_selectpage = "SELECT Page_name FROM Store_pages WHERE Store_ID =" & store_id & " and page_id= "&pageselect 
	rs_store.open str_selectpage,conn_store,1,1
	varpagename= rs_store("Page_name")
	response.write "state is " &  rs_store.state 

	'ADD FIELD TO STORE_CUSTOM_FIELDS TABLE
	
	if request.form("op") <> "edit" then
	
		sql_add = " insert into store_forms (store_id,fldpageid,fldpagename,fldToEmail,fldsubjectform)"
		sql_add = sql_add & " values("&store_id & ",'" &pageselect&"','" & varpagename&  "','"&Emailto&"','"&Subject&"')"
		Response.Write "<BR>sql_add=" & sql_add
		conn_store.Execute sql_add
	else
		sql_add = " Update store_forms set fldpageid ='" &pageselect&"',fldpagename='"&varpagename&"',fldToEmail='"&Emailto&"', fldsubjectform ='"&Subject&"' where store_id=" & store_id & " and fldpbid="& varid 
		conn_store.Execute sql_add
		
		call generateForm(varid)

	end if
  end if
	
	       Response.Redirect "fields_page.asp"

end if
%>
