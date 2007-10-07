<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "marketing/affiliates.htm"
sTextHelp="affiliates/affiliates.doc"


set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Affiliates"
myStructure("ColumnList") = "affiliate_id,code,contact_name,email,url"
myStructure("HeaderList") = "code,contact_name,email,url,reports"
myStructure("DefaultSort") = "code"
myStructure("PrimaryKey") = "affiliate_id"
myStructure("Level") = 7
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "affiliates_Manager.asp"
myStructure("FormAction") = "affiliates_Manager.asp"
myStructure("Title") = "Affiliates"
myStructure("FullTitle") = "Marketing > Affiliates"
myStructure("CommonName") = "Affiliate"
myStructure("NewRecord") = "new_affiliate.asp"
myStructure("Heading:affiliate_id") = "PK"
myStructure("Heading:code") = "Name"
myStructure("Heading:contact_name") = "Contact"
myStructure("Heading:email") = "Email"
myStructure("Heading:url") = "URL"
myStructure("Heading:reports") = "Reports"
myStructure("Format:code") = "STRING"
myStructure("Format:contact_name") = "STRING"
myStructure("Format:email") = "STRING"
myStructure("Format:url") = "STRING"
myStructure("Format:reports") = "TEXT"
myStructure("Link:reports") = "affiliates_reports.asp?aff_id=PK"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
	end if
	createFoot thisRedirect, 0
%>

