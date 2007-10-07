<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Australian Merchant Credit Card Processing"
thisRedirect = "ausmerchants.asp"
sMenu="none"
createHead thisRedirect  %>


			  </form>
		<TR><td colspan=2><HR></td></tr>
			 <tr><td><UL><LI>AU$200 Setup Fee
			 <LI>AU$430 Annual Payment Processing Fee (3 currencies)
			 <LI>4.5% Discount Rate
			 <LI>Includes merchant account
			 </UL></td>
			 <td><form action="https://secure.worldpay.com/app/splash.pl?Pid=62347" method="Post">
<input type=submit name="WorldPay" value="Apply">
<BR><BR></form></td></tr>
		<TR><TD colspan=2><HR></td></tr>
			 <tr><td>
			 <UL><LI>AU$99.00 Application Fee
			 <LI>AU$35.00 Monthly Fee
			 <LI>50¢ transaction fee
			 <LI>Need separate merchant account
			 </UL></td>
			 <td><form action=http://www.eway.com.au><input type=submit name=eWay value=Apply></form></td></tr>
			 <TR><TD colspan=2><HR></td></tr>
			 <tr><td>
			 <UL><LI>No Application Fee
			 <LI>AU$285.00 Yearly Fee
			 <LI>50¢ transaction fee
			 <LI>Need separate merchant account
			 </UL></td>
			 <td><form action=http://www.eway.com.au><input type=submit name=eWay value=Apply></form></td></tr>
			 <TR><TD colspan=2><HR></td></tr>
			 <tr><td>
			 <UL><LI>$49 Application Fee
			 <LI>$0 Monthly Fee
			 <LI>5.5% Discount Rate
			 <LI>45¢ transaction fee
			 <LI>Customers will be redirected on checkout to 3rd party site for payment processing
			 <LI>High Risk Merchants Accepted
			 <LI>Credit card statement will NOT show merchants name
			 </UL></td>
			 <td><form action=http://www.2checkout.com/cgi-bin/aff.2c?affid=74224><input type=submit name=2Checkout value=Apply></form></td></tr>

	  <TR><TD></td></tr>

<% createFoot thisRedirect, 0 %>


