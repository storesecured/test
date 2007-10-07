<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sql_region = "SELECT Country, Country_Id FROM Sys_Countries ORDER BY Country;"
set myfields1=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_region,mydata1,myfields1,noRecords1)
sQuestion_Path = "advanced/customer_groups.htm"
sFlashHelp="customergroups/customergroups.htm"
sMediaHelp="customergroups/customergroups.wmv"
sZipHelp="customergroups/customergroups.zip"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Customers_Groups"
myStructure("ColumnList") = "group_id,group_name,group_country,group_budget_min,group_purchase_history"
myStructure("DefaultSort") = "group_name"
myStructure("PrimaryKey") = "group_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "customers"
myStructure("FileName") = "customers_groups.asp"
myStructure("FormAction") = "customers_groups"
myStructure("Title") = "Customer Groups"
myStructure("FullTitle") = "Customers > Groups"
myStructure("CommonName") = "Customer Group"
myStructure("NewRecord") = "Create_New_Customer_Group.asp"
myStructure("Heading:group_id") = "PK"
myStructure("Heading:group_name") = "Name"
myStructure("Heading:group_country") = "Country"
myStructure("Heading:group_budget_min") = "Budget min"
myStructure("Heading:group_purchase_history") = "Purchase Min"
myStructure("Format:group_name") = "STRING"
myStructure("Format:group_country") = "COLLECTION"
myStructure("Format:group_budget_min") = "CURR"
myStructure("Format:group_purchase_history") = "CURR"
myStructure("Collection:group_country") = "country,country"
myStructure("Default:group_country") = "All Countries"


%>
<!--#include file="head_view.asp"-->
<!-- start of list view -->
<!--#include file="list_view.asp"-->
<!-- end of list view -->

<%
	if Request.QueryString("Delete_Id") <> "" then

	end if
createFoot thisRedirect, 0
%>

