<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Merchant Accounts"
thisRedirect = "merchants.asp"
sMenu="none"
createHead thisRedirect  %>

			  <TR><td colspan=2>Choose your location from the list below to see Easystorecreator's recommended credit card processing solutions.

		<UL><LI><a class=link href=usmerchants.asp>US Based Merchants</a>
			 <LI><a class=link href=cdnmerchants.asp>Canadian Based Merchants</a>
			 <LI><a class=link href=ukmerchants.asp>UK Based Merchants</a>
			 <LI><a class=link href=intlmerchants.asp>International Merchants</a>
			 </UL></td></tr>
			 <TR><td colspan=2><B>Common terms defined</b>
			 <UL><LI>Discount Rate-Percentage of each transaction which is charged by the processing bank
			 <LI>Transaction Fee-Flat fee charged for each transaction processed
			 <LI>Monthly Fee-Flat monthly charge to process credit cards
			 <LI>Application Fee-Charge to setup credit card processing account.  Usually refundable if account denied for any reason
			 </UL>


			 </td></tr>

<% createFoot thisRedirect, 0 %>


