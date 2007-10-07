<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
sQuestion_Path = "import_export/my_account/feature_request.htm"
sInstructions="Please note that feature requests are not meant for submitting questions.  This area is a spot for you to enter new functionality you would like to see in upcoming releases.  Feature Requests are not responded to they are reviewed by the development team approximately once a month when deciding what new features to implement."
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "sys_features"
myStructure("ColumnList") = "feature_id,sys_created,subject,detail"
myStructure("HeaderList") = "sys_created,subject,detail"
myStructure("DefaultSort") = "sys_created"
myStructure("PrimaryKey") = "feature_id"
myStructure("Level") = 3
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "account"
myStructure("FileName") = "features2.asp"
myStructure("FormAction") = "features2.asp"
myStructure("Title") = "Feature Requests"
myStructure("FullTitle") = "My Account > Feature Requests"
myStructure("CommonName") = "Feature"
myStructure("NewRecord") = "feature_request.asp"
myStructure("Heading:feature_id") = "PK"
myStructure("Heading:sys_created") = "Request Date"
myStructure("Heading:subject") = "Subject"
myStructure("Heading:detail") = "Detail"
myStructure("Heading:status") = "Status"
myStructure("Heading:area") = "Area"
myStructure("Heading:area2") = "Section"
myStructure("Format:sys_created") = "DATESHORT"
myStructure("Format:subject") = "STRING"
myStructure("Format:area") = "STRING"
myStructure("Format:detail") = "STRING"
myStructure("Format:status") = "STRING"
myStructure("Format:area2") = "STRING"
%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->


<% createFoot thisRedirect, 0%>

