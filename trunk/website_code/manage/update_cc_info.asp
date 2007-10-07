<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include virtual="common/crypt.asp"-->

<% 

'RETRIVE CURRENT VALUES FROM THE DATABASE
sql_select = "select * from Sys_Billing where Store_Id = "&Store_Id

rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	First_name= rs_store("First_name")
	Last_name= rs_store("Last_name")
	Company= rs_store("Company")
	Address= rs_store("Address")
	City= rs_store("City")
	State= rs_store("State")
	Zip= rs_store("Zip")
	Country= rs_store("Country")
	Phone= rs_store("Phone")
	Fax= rs_store("Fax")
	EMail= rs_store("EMail")
	Card_Number=rs_store("Card_Number")
	Exp_Month=rs_Store("Exp_Month")
	Exp_Year=rs_Store("Exp_Year")
	Payment_Method_Current = replace(rs_Store("Payment_Method")," ","")
	Payment_Term = rs_Store("Payment_Term")
	Bank_Name = rs_Store("Bank_Name")
	Bank_Account = decrypt(rs_store("Bank_Account"))
	Bank_ABA = decrypt(rs_store("Bank_ABA"))
	Acct_Type = rs_Store("Acct_Type")
	Org_Type = rs_Store("Org_Type")
	Check_Num = rs_Store("Check_Num")
	License_Num = rs_Store("License_Num")
	License_State = rs_Store("License_State")
	License_Exp_Day = rs_Store("License_Exp_Day")
	License_Exp_Month = rs_Store("License_Exp_Month")
	License_Exp_Year = rs_Store("License_Exp_Year")
	reseller_portion = rs_store("Reseller_Portion")
else

	Company= Store_Company
	Address= Store_Address1
	City= Store_City
	Zip= Store_Zip
	Country= Store_Country
	Phone= Store_Phone
	Fax= Store_Fax
	EMail= Store_Email
	reseller_portion = 0
end if
rs_store.close

sFormAction = "process_billing.asp"
sFormName = "payment"
sTitle = "Update Payment Method Continue"
sFullTitle = "My Account > Payments > <a href=update_cc.asp class=white>Update Payment Method</a> > Continue"
thisRedirect = "billing_info.asp"
sMenu = "account"
addPicker = 1
sSubmitName = "submit"
createHead thisRedirect

payment_method = request.form("payment_method")
term = request.form("term")

on error resume next

if payment_method = "Paypal" or payment_method = "PayPal" then %>
  <% if Payment_Method_Current = "Paypal" or Payment_Method_Current = "PayPal" then %>
  <tr bgcolor='#FFFFFF'><TD>Since you have selected Paypal as your method of payment you must login to your Paypal
  account to make any changes to your billing information.
  <BR><BR><a href=http://www.paypal.com class=link target=_blank>Click here to login to PayPal</a>
  </TD></TR>
  <% else %>
  <tr bgcolor='#FFFFFF'><TD>To switch to Paypal please click here
  </form><form method="POST"	action=billing_info.asp name=>
  <input type=hidden name=Payment_Method value="Paypal">
  <% if Service_Type = 3 then
		 sLevel = "bronze"
	elseif Service_Type = 5 then
		 sLevel = "silver"
	elseif Service_Type = 7 then
		 sLevel = "gold"
	elseif Service_Type = 9 then
		 sLevel = "platinum"
	elseif Service_Type = 5 then
		 sLevel = "unlimited"
	end if %>
  <input type=hidden name=Service value="<%= sLevel %>">
  <input type=hidden name=term value="<%= Payment_Term %>">
  <input type=hidden name=Type value="Normal">
  <input type=hidden name=hidResellerAmt value="<%= reseller_portion %>">
  <input type=submit name="Submit" value="Switch to Paypal">
  </form>
  </TD></TR>
  <% end if %>
<BR>
<input type=hidden name=hidResellerAmt value="<%= reseller_portion %>">

<!--#include file="capture_billing.asp"-->
<%
else
  if Payment_Method_Current = "Paypal" or Payment_Method_Current = "PayPal" then %>
  <tr bgcolor='#FFFFFF'><TD colspan=3>After successfully updating your payment method to <%= Payment_Method %>, please cancel your previous Paypal subscription to avoid double billing.
  <BR><BR></TD></TR>
  <% end if

	 put_normal
%>

    <input type=hidden name=hidResellerAmt value="<%= reseller_portion %>">

	 <tr bgcolor='#FFFFFF'><TD colspan=3 align=center><input type=hidden name=billing_type value="Update">
<input type="submit" value="Save" name=<%= sSubmitName %>></td></tr>
<% end if %>

<%   createFoot thisRedirect,2
if payment_method <> "Paypal" and payment_method <> "PayPal" then %>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("first","req","Please enter your first name.");
 frmvalidator.addValidation("last","req","Please enter your last name.");
 frmvalidator.addValidation("address","req","Please enter your address");
 frmvalidator.addValidation("city","req","Please enter your city");
 frmvalidator.addValidation("State","req","Please enter your state.");
 frmvalidator.addValidation("zip","req","Please enter your zip.");
 frmvalidator.addValidation("Phone","req","Please enter your phone number.");
 frmvalidator.addValidation("email","req","Please enter your email address.");
 frmvalidator.addValidation("email","email","Please enter your email address.");
<% if payment_method = "eCheck" then %>
frmvalidator.addValidation("BankName","req","Please enter your bank name.");
frmvalidator.addValidation("BankABA","req","Please enter your bank routing number.");
frmvalidator.addValidation("BankAccount","req","Please enter your bank account number.");
frmvalidator.addValidation("CheckSerial","req","Please enter your check serial number.");
frmvalidator.addValidation("DrvNumber","req","Please enter your drivers license number.");
frmvalidator.addValidation("DrvState","req","Please enter your drivers license state.");
frmvalidator.addValidation("dobd","req","Please enter your drivers license expiration.");
frmvalidator.addValidation("dobm","req","Please enter your drivers license expiration.");
frmvalidator.addValidation("doby","req","Please enter your drivers license expiration.");
<% end if %>
</script>
<%
end if

%>


