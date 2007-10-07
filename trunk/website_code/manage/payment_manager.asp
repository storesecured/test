<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sInstructions="Check the methods of payment which you will accept.  At checkout time your customers will be able to choose from all payment methods which you have enabled below.<BR><BR>Visa, Mastercard, American Express, Discover and Diners Club are reserved words and are used for credit card processing.<BR><BR>Paypal is a reserved payment method name associated with their respective payment processor.<BR>eCheck is a reserved word used for electronic check processing by the few providers who accept this method of payment.<BR><BR>Charge my account and Free payment methods are reserved words with special internal meaning for customers with budgets and free orders at checkout respectively.<BR><BR>Debit Card is a special reserved word for certain UK payment processors who treat Debit Cards separately from credit cards.  Note US payment processors do not use this.<BR><BR>All other payment method names are simply recorded with no special action taken at checkout."
sTextHelp="paymentmethods/payment_methods.doc"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_payment_methods"
myStructure("ColumnList") = "store_payment_id,payment_name,accept"
myStructure("HeaderList") = "store_payment_id,payment_name,accept"
myStructure("DefaultSort") = "payment_name"
myStructure("PrimaryKey") = "store_payment_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "general"
myStructure("FileName") = "payment_manager.asp"
myStructure("FormAction") = "payment_manager.asp"
myStructure("Title") = "Payment Methods"
myStructure("FullTitle") = "General > Payments > Methods"
myStructure("CommonName") = "Payment Method"
myStructure("NewRecord") = "payment_edit.asp"
myStructure("Heading:store_payment_id") = "PK"
myStructure("Heading:payment_name") = "Name"
myStructure("Heading:accept") = "Enabled"
myStructure("Format:payment_name") = "STRING"
myStructure("Format:accept") = "LOOKUP"
myStructure("Lookup:accept") = "0:No^-1:Yes"
%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
createFoot thisRedirect, 0
%>

