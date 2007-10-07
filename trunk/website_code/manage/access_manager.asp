<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sInstructions ="The log below will show all logins to your store admin area along with the username and ip address used.  This log can help you monitor who is logging in to manage the store and when."
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_access_log"
myStructure("ColumnList") = "sys_created,clientip,user_name"
myStructure("HeaderList") = "sys_created,clientip,user_name"
myStructure("DefaultSort") = "sys_created"
myStructure("PrimaryKey") = "access_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 0
myStructure("BackTo") = ""
myStructure("Menu") = "account"
myStructure("FileName") = "access_manager.asp"
myStructure("FormAction") = "access_manager.asp"
myStructure("Title") = "Admin Login Log"
myStructure("FullTitle") = "My Account > <a href=security.asp class=white>Admin Logins</a> > Logs"
myStructure("CommonName") = "Admin Login Log"
myStructure("NewRecord") = ""
myStructure("Heading:access_id") = "PK"
myStructure("Heading:sys_created") = "Date"
myStructure("Heading:clientip") = "IP Address"
myStructure("Heading:user_name") = "User Name"
myStructure("Format:sys_created") = "DATE"
myStructure("Format:clientip") = "STRING"
myStructure("Format:user_name") = "STRING"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%

createFoot thisRedirect, 0
%>

