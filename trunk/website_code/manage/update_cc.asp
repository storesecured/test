<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
if checkSSL=1 then
   scheckSSL=1
else
   scheckSSL=0
end if

sql_select = "select Payment_Method,Payment_Term, Amount from sys_billing where store_id="&Store_Id
rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	Payment_Term = rs_Store("Payment_Term")
	Payment_Method = replace(rs_Store("Payment_Method")," ","")
	Amount = rs_Store("Amount")
end if
rs_store.close

if scheckSSL=1 then
   sFormAction = "https://manage.storesecured.com/update_cc_info.asp"
   sFormAction = "update_cc_info.asp"
else
   sFormAction = "update_cc_info.asp"
end if

sFormName = "payment"
sTitle = "Update Payment Method"
sFullTitle = "My Account > Payments > Update Payment Method"
thisRedirect = "billing_info.asp"
sMenu = "account"
sSubmitName="submit"
sQuestion_Path = "import_export/my_account/update_payment.htm"
createHead thisRedirect

if Trial_Version then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available for trial accounts.<BR><BR>
		<% if Trial_Version then %>
		<a href=https://manage.storesecured.com/billing.asp class=link>Click here to upgrade now.</a>
		<% else %>
		<a href=https://manage.storesecured.com/update_cc.asp class=link>Click here to upgrade now.</a>
		<% end if %>
	</td></tr>

<% else %>


	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Type</B></TD><TD class="inputvalue"><select name="payment_method" size="1">
			 <% if Payment_Method <> "" then %>
			 <option selected value="<%= Payment_Method %>"><%= Payment_Method %></option>
			 <% end if %>
			 <option value="Visa">Visa</option>
			 <option value="Mastercard">Mastercard</option>
			 <option value="American Express">American Express</option>
			 <option value="Discover">Discover</option>
			 <option value="Paypal">Paypal</option>
			 </select></td><td class=inputvalue><img src=images/icon_visa.gif><img src=images/icon_mastercard.gif><img src=images/icon_discover.gif><img src=images/icon_amex.gif><BR>OR <B>PayPal</b>


	 <% small_help "Type" %></TD>
	 </TR>

	 <TR bgcolor='#FFFFFF'><TD colspan=4 align=center><input type="submit" value="Continue" name=<%= sSubmitName %>></td></tr>
<% end if %>
<% if scheckSSL=0 then %>
      <table bgcolor=Red width=100%><TR bgcolor='#FFFFFF'><td class=error><B>WARNING: Your browser does not support secure pages</b><BR>We are unable to transfer you to a secure connection to collect your sensitive payment information because either your browser does not support a secure connection or you have disabled it.  You may want to choose a new browser or enable SSL before continuing.  You may also call our toll-free number 1-866-324-2764 to upgrade by phone.</b></td></tr></table>
<% end if %>

<% createFoot thisRedirect,2 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>

