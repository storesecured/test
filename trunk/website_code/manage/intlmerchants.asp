<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "International Merchant Credit Card Processing"
thisRedirect = "intlmerchants.asp"
sMenu="none"
createHead thisRedirect  %>


			  </form>
			  <TR><td colspan=2><HR></td></tr>
			 <tr><td><UL><LI>Processor rates are variable depending on country.	Click on the apply link to select a country and find the applicable rates.
			 </UL></td>
			 <td><form action="https://secure.worldpay.com/app/splash.pl?Pid=62347" method="Post">
<input type=submit name="WorldPay" value="Apply">
<BR><BR></form></td></tr>

		<TR><TD colspan=2><HR></td></tr>
			 <tr><td>
			 <UL><LI>$49 Application Fee
			 <LI>$0 Monthly Fee
			 <LI>5.5% Discount Rate
			 <LI>45 cent transaction fee
			 <LI>Instant Approval
			 <LI>Customers will be redirected on checkout to 3rd party site for payment processing
			 <LI>Customers billing statment will not show merchants name
			 <LI>High Risk Merchants Accepted
			 </UL></td>
			 <td><form action=http://www.2checkout.com/cgi-bin/aff.2c?affid=74224><input type=submit name=2Checkout value=Apply></form></td></tr>

	  <TR><TD></td></tr>

<% createFoot thisRedirect, 0 %>


