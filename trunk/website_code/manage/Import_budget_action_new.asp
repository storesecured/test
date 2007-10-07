<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Budget"
sHeader="budget"
sMenu="customers"
sTitle="Customer Budget Import" 
sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > Budget > Import"

Dim arrColumns(2)
arrColumns(0) = "Budget_Amount:Re:Integer:"
arrColumns(1) = "Customer_Id:Re:Integer:"
arrColumns(2) = "Notes:Op:Text:1000"

sProcName = "wsp_budget_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
