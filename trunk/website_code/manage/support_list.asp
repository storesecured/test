<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

if request.querystring("Support_Id") <> "" and request.querystring("Status") <> "" then
	sql_update = "update sys_support set Status="&request.querystring("Status")&" where support_Id="&request.querystring("Support_Id")&" and store_id="&Store_Id
	conn_store.Execute sql_update


	sql_select="select top 1 Created_By from Sys_Support_Detail where  Support_Id="&request.querystring("Support_Id")
	set rs_name = conn_store.Execute(sql_select)
	
    Detail=""
    Name=rs_name("created_by")


   	IP_Address= Request.ServerVariables("REMOTE_ADDR")

	sql_insert="New_Support_Details1 "&request.querystring("Support_Id")&",'"&Detail&"','"&Name&"','"&IP_Address&"',"&request.querystring("Status")
	conn_store.Execute sql_insert
end if


sFormAction = "support_list.asp"
sTitle = "My Account > Support Requests > View/Edit"
thisRedirect = "support_list.asp"
sMenu = "support"
sQuestion_Path = "import_export/my_account/support_requests.htm"
sInstructions="Editing a support request with new information will reopen it and resend it to the support team."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "sys_support"
myStructure("ColumnList") = "support_id,sys_created,subject,status"
myStructure("HeaderList") = "sys_created,subject,status"
myStructure("DefaultSort") = "sys_created"
myStructure("PrimaryKey") = "support_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "account"
myStructure("FileName") = "support_list.asp"
myStructure("FormAction") = "support_list.asp"
myStructure("Title") = "Support Requests"
myStructure("FullTitle") = "My Account > Support Requests"
myStructure("CommonName") = "Support Request"
myStructure("NewRecord") = "support_request.asp"
myStructure("EditRecord") = "support_edit.asp"
myStructure("Heading:support_id") = "PK"
myStructure("Heading:sys_created") = "Request Date"
myStructure("Heading:subject") = "Subject"
myStructure("Heading:status") = "Status"
myStructure("Format:sys_created") = "DATESHORT"
myStructure("Format:subject") = "STRING"
myStructure("Format:status") = "LOOKUP"
myStructure("Lookup:status") = "1:New^2:Assigned^3:Cancelled^4:Completed^5:Reopened^6:Awaiting Feedback"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->
<%
Delete_Id = fn_get_delete_ids
if Delete_Id<>"" then
    sql_delete = "delete from sys_support_detail where support_id in ("&delete_id&")"
	conn_store.Execute sql_delete
end if
%>
<% createFoot thisRedirect, 0%>

