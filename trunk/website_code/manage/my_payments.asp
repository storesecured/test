<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sql_select = "select Sys_Billing.Fee_Discount, Sys_Billing.Discount_Until, Sys_Billing.Amount, Sys_Billing.next_billing_date from Sys_Billing where sys_billing.store_id="&Store_Id
rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	Amount= rs_store("Amount")
	Fee_Discount = rs_store("Fee_Discount")
	Discount_Until = rs_store("Discount_Until")
	next_billing_date = rs_Store("next_billing_date")
end if
rs_store.close

if Service_Type = 3 then
	Service = "Bronze"
elseif Service_Type = 5 then
	Service = "Silver"
elseif Service_Type = 7 then
	Service = "Gold"
elseif Service_Type = 9 or Service_Type = 10 then
	Service = "Platinum"
elseif Service_Type = 11 then
	Service = "Unlimited"

else
	Service = "Unknown"
end if

if Discount_Until > Now() then
	Amount = Amount - Fee_Discount
end if

  if Discount_Until > BillDate then
	  Amount = Amount - Fee_Discount
  end if
  if Amount > 0 then
		sPaymentInfo = "Your next payment of <B>$"&Amount&"</b> for "&Service&" Service will be automatically billed on <B>"&formatdatetime(next_billing_date,2)&"</b>"
  end if

sQuestion_Path = "import_export/my_account/prior_invoices.htm"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "sys_payments"
myStructure("ColumnList") = "payment_id,transdate,amount,payment_description"
myStructure("HeaderList") = "transdate,amount,payment_description,details"
myStructure("SortList") = "transdate,amount,payment_description"
myStructure("DefaultSort") = "transdate"
myStructure("PrimaryKey") = "payment_id"
myStructure("Level") = 1
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 0
myStructure("BackTo") = ""
myStructure("Menu") = "account"
myStructure("FileName") = "my_payments.asp"
myStructure("FormAction") = "my_payments.asp"
myStructure("Title") = "Prior Invoices"
myStructure("FullTitle") = "My Account > Payments > Prior Invoices"
myStructure("CommonName") = "Payment"
myStructure("NewRecord") = "my_invoice.asp"
myStructure("Heading:payment_id") = "PK"
myStructure("Heading:transdate") = "Date"
myStructure("Heading:amount") = "Amount"
myStructure("Heading:payment_description") = "Description"
myStructure("Heading:details") = "Details"
myStructure("Format:transdate") = "DATESHORT"
myStructure("Format:amount") = "CURR"
myStructure("Format:payment_description") = "STRING"
myStructure("Format:details") = "TEXT"
myStructure("Link:details") = "my_invoice.asp?Payment_Id=PK"
%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

	<TR bgcolor='#FFFFFF'><td><BR><%= sPaymentInfo %></td></tr>

<% createFoot thisRedirect, 0%>
