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
	reseller_portion = rs_store("reseller_portion")
else

	Company= Store_Company
	Address= Store_Address1
	City= Store_City
	Zip= Store_Zip
	Country= Store_Country
	Phone= Store_Phone
	Fax= Store_Fax
	EMail= Store_Email
end if
rs_store.close

sFormAction = "process_billing.asp"
sFormName = "payment"
sTitle = "Billing Info"
thisRedirect = "billing_info.asp"
sMenu = "account"
addPicker = 1
sSubmitName ="submit"
createHead thisRedirect

payment_method = request.form("payment_method")
term = request.form("term")
sBillingType="Custom"

if Custom_Description="Manual Payment" then
   hidResellerAmt = reseller_portion
   sBillingType="Normal"
end if

on error resume next
if payment_method = "Paypal" or payment_method="PayPal" then
	Return_Addr = "http://"&Request.ServerVariables("HTTP_HOST")&"/old_billing_info.asp?Accepted=Yes&Billing_Type="&sBillingType
	Notify_Addr = "http://"&Request.ServerVariables("HTTP_HOST")&"/paypal_notify1.asp"
	Cancel_Addr = "http://"&Request.ServerVariables("HTTP_HOST")&"/paypal_cancel.asp"

	%>
	</form>
	<form method='POST' action='https://www.paypal.com/cgi-bin/webscr' name='Payment'>
	<input type='hidden' name='business' value='<%=sPaypal_email %>'>
	<input type="hidden" name="cmd" value="_ext-enter">
	<input type="hidden" name="redirect_cmd" value="_xclick">
	<input type="hidden" name="bn" value="EasyStoreCreator.EasyStoreCreator">
	<TR bgcolor='#FFFFFF'><td colspan=3>Click on the button below to pay using Paypal</td></tr>
	<TR bgcolor='#FFFFFF'><td colspan=3><BR><BR><B>Important:</b> Please note that payments made using <B>Paypals eCheck</b> service are not recorded until the payment has cleared.<BR></td></tr>
	 <TR bgcolor='#FFFFFF'><td colspan=3><BR></td></tr>
	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Amount Due</B></TD><TD class="inputvalue">$<%= Custom_Amount %>

	 <% small_help "Charge Total" %></TD>
	 </TR>
	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Description</B></TD><TD class="inputvalue"><%= Custom_Description %>

	 <% small_help "Charge Total" %></TD>
	 </TR>
	 

	<TR bgcolor='#FFFFFF'><td colspan=3 align=center>

	<input type="hidden" name="item_number" value="<%= Store_Id %>">

	<INPUT type="hidden" NAME="amount" value="<%= Custom_Amount %>">
	<INPUT type="hidden" NAME="item_name" value="<%= Custom_Description %>">
	<input type="hidden" value="<%=hidResellerAmt%>" name="hidResellerAmt">
	<input type="hidden" name="notify_url" value="<%= Notify_Addr %>?Level=<%= level %>&Term=<%= term %>&Term_Name=<%= term_name %>&Service=<%= server.urlencode(Custom_Description) %>&Type=Custom&Grand_Total=<%= Custom_Amount %>">

<!--#include file="capture_billing.asp"-->
<% put_paypal %>

<%
else
	put_normal

  IPAddress = Request.ServerVariables("REMOTE_ADDR")
  Set ipObj = Server.CreateObject("IP2Location.Country")
  ' initialize IP2Location™ Component
  If ipObj.Initialize("melanie@easystorecreator.com-2005-7-pQ3GpeMWWo1m") <> "OK" Then
  'initialization failed
  End If
  ' it will return "US"
  CountryName = ipObj.LookUpFullName(IPAddress)
  Set ipObj = nothing
%>
	<TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Amount Due</B></TD><TD class="inputvalue">$<%= Custom_Amount %>
	 <% small_help "Charge Total" %></TD>
	 </TR>
	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><input type="checkbox" name="Save_Info" value="TRUE" checked></TD><TD class="inputvalue"><B>Save Information for future payments</b>

	 <% small_help "Save Information" %></TD>
	 </TR>
	 <input type=hidden name=total value="<%= Custom_Amount %>">
	 <input type=hidden name=counter value="<%= DatePart("d", Date()) %><%= Store_Id %>">
	 <input type=hidden name=term value="<%= term %>">
	 <input type=hidden name=term_name value="<%= term_name %>">
	 <input type=hidden name=service value="<%= service %>">
	 <input type=hidden name=level value="<%= level %>">
	 <input type=hidden name=description value="<%=Custom_Description %>">
	 <input type=hidden name=ponum value="<%= Store_Id %>-<%= datepart("m",date())%>/<%= datepart("yyyy",date())%>">
	 <input type=hidden name=billing_type value="Custom">
	 <input type="hidden" value="<%= hidResellerAmt %>" name="hidResellerAmt">
	 <TR bgcolor='#FFFFFF'><TD colspan=3 align=center><input type="submit" value="Pay Now" name=<%=sSubmitName %>></td></tr>
   <TR bgcolor='#FFFFFF'><TD colspan=3 align=center><input type=hidden name=ip_address value=<%= IPAddress %>><input type=hidden name=ip_country value="<%= CountryName %>">
   <B>Charges will show on your statement as <I>StoreSecured, Inc.</i></b><BR><BR>
   Your IP Address (<%= IPAddress %>) has been recorded and will be provided to the necessary authorities in the event of unauthorized or fraudulent charges.</td></tr>
   

<%

end if

createFoot thisRedirect,2

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
<% end if %>


