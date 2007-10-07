<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "UK Merchant Credit Card Processing"
thisRedirect = "ukmerchants.asp"
sMenu="none"
createHead thisRedirect  %>


			  </form>
		<TR bgcolor="#FFFFFF"><td colspan=2><HR></td></tr>
			 <tr bgcolor="#FFFFFF"><td><UL><LI>£100 Setup Fee
			 <LI>£160 Annual Fee
			 <LI>4.5% Discount Rate
			 <LI>Includes Merchant Account
			 </UL></td>
			 <td><form action="https://secure.worldpay.com/app/splash.pl?Pid=62347" method="Post">
<input type=submit name="PSIGate" value="Apply">
<BR><BR></form></td></tr>

	  
<% createFoot thisRedirect, 0 %>


