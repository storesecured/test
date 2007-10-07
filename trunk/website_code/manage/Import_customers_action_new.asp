<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Customers"
sTitle = "Import Customers"
sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > Import"
sMenu="customers"
sHeader="User"

Dim arrColumns(28)

arrColumns(0) = "User_id:Re:Text:50::--ExpressCheckOut Customer--"
arrColumns(1) = "Password:Re:Text:50"
arrColumns(2) = "Last_Name:Re:Text:100"
arrColumns(3) = "First_Name:Re:Text:100"
arrColumns(4) = "Company:Op:Text:100"
arrColumns(5) = "Address1:Re:Text:200"
arrColumns(6) = "Address2:Op:Text:200"
arrColumns(7) = "City:Re:Text:200"
arrColumns(8) = "State:Re:Text:50"
arrColumns(9) = "Zip_Code:Re:Text:50"
arrColumns(10) = "Country:Re:Text:200"
arrColumns(11) = "Phone:Re:Text:100"
arrColumns(12) = "Fax:Op:Text:50"
arrColumns(13) = "Email_Address:Re:Text:100"
arrColumns(14) = "Last_Name:Re:Text:100:%arrColumns%2"
arrColumns(15) = "First_Name:Re:Text:100:%arrColumns%3"
arrColumns(16) = "Company:Op:Text:100:%arrColumns%4"
arrColumns(17) = "Address1:Re:Text:200:%arrColumns%5"
arrColumns(18) = "Address2:Op:Text:200:%arrColumns%6"
arrColumns(19) = "City:Re:Text:200:%arrColumns%7"
arrColumns(20) = "State:Re:Text:50:%arrColumns%8"
arrColumns(21) = "Zip_Code:Re:Text:50:%arrColumns%9"
arrColumns(22) = "Country:Re:Text:200:%arrColumns%10"
arrColumns(23) = "Phone:Re:Text:100:%arrColumns%11"
arrColumns(24) = "Fax:Op:Text:50:%arrColumns%12"
arrColumns(25) = "Email_Address:Re:Text:100:%arrColumns%13"
arrColumns(26) = "Send_Newsletter:Re:Boolean:"
arrColumns(27) = "Tax_Exempt:Re:Boolean:"
arrColumns(28) = "Protected_Access:Re:Boolean:"
sProcName = "wsp_customers_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
