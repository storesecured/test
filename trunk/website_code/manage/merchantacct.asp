<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
if Affiliate_Id = "35" then
   response.redirect "http://www.meracard.com"
end if

sTitle = "Accept Credit Cards"
sFullTitle = "General > Payments > Accept Credit Cards"
thisRedirect = "merchantacct.asp"
sFormAction = "merchantacct.asp"
sMenu="general"
createHead thisRedirect  

%>

   <% if Store_Country = "United States" then %>
<tr bgcolor='#FFFFFF'><td>
    <B>US Customers</B><BR>

All plans listed below include:<BR>
<table><tr bgcolor='#FFFFFF'>
<td valign=top>
<UL><LI>NO Application Fee
<LI>Quick Approval
</UL>
</td>
<td width=40></td>
<td valign=top>
<UL><LI>NO Setup Fee
<LI>NO Annual Fee
</UL>
</td></tr></table>

<BR>
<table border=1 cellspacing=0 cellpadding=5 width="100%">
<tr bgcolor=DDDDDD><td><B>Plan</b></td><td><b>Discount Rate</b></td><td><b>Transaction Fee</b></td><td><b>Monthly Fee</b></td><td><b>Monthly Minimum</b></td><td><b>Recommended for merchants who have...</b></td><td><b>Apply Now</td></tr>
<tr bgcolor='#FFFFFF'><td>L-Business</td>
<td>2.19%</td><td>28¢</td><td>$25</td><td>$25</td><td>monthly sales > $2000</td><td><a class=link href=https://www.bluemerchant.com/bluepay/merchantform/form0.asp?hostID=506302 target=_blank>Apply Online</a></tr>
<tr bgcolor=DDDDDD><td>M-Small Business</td>
<td>2.49%</td><td>35¢</td><td>$20</td><td>$25</td><td>monthly sales $1000 - $2000</td><td><a class=link href=https://www.bluemerchant.com/bluepay/merchantform/form0.asp?hostID=506301 target=_blank>Apply Online</a></tr>
<tr bgcolor=#FFFFFF><td>S-Starter</td>
<td>2.99%</td><td>50¢</td><td>$15</td><td>$25</td><td>monthly sales < $1000</td><td><a class=link href=https://www.bluemerchant.com/bluepay/merchantform/form0.asp?hostID=504911 target=_blank>Apply Online</a></tr>
<tr bgcolor='#FFFFFF'><td colspan=8>For customers who cannot get approved for a normal merchant account click here to find out more about <a class=link href=http://www.easymerchantservices.com/third-party-merchant-accounts.asp>3rd party merchant accounts</a>.</td></tr>

</table>




</td></tr>

    <tr bgcolor='#FFFFFF'><td><BR>





    </td></tr>
    <% elseif Store_Country = "Canada" then %>
    <tr bgcolor='#FFFFFF'><td><CENTER><B>Recommended package #1 for Canadian high volume stores</b></CENTER>
	<UL><LI>$199 CDN Application Fee
<LI>$49.95 CDN Monthly Fee
<LI>3.75% Discount Rate
<LI>35¢ CDN Transaction Fee
</UL>
	<center><a class=link href=https://www.psigate.com/wizard/page0.asp?partner=1102703 target=_blank class=link
onMouseOver="status='http://www.psigate.com'; return true" onMouseOut="status=''; return true"><B>Apply now</b></a></center>
	</td>
	<td><CENTER><B>Recommended package #2 for Canadian low volume stores</b></CENTER>
	<UL><LI>$99 CDN Application Fee
<LI>No Monthly Fee
<LI>9% Discount Rate
<LI>No Transaction Fee
</UL>
	<center><a class=link href=https://www.internetsecure.com/newapplication.asp?ReferID=745 target=_blank class=link
onMouseOver="status='http://www.internetsecure.com'; return true" onMouseOut="status=''; return true"><B>Apply now</b></a></center>
	</td>
	</tr>
	<tr bgcolor='#FFFFFF'><td colspan=2 align=center><font size=1><a href=cdnmerchants.asp class=link><font size=1>Click here to see other Canadian options</font></a>
	<% elseif Store_Country = "United Kingdom" then %>
	<tr bgcolor='#FFFFFF'><td><CENTER><B>Recommended package #1 for UK stores</b></CENTER>
	<UL><LI>4.5% Discount Rate
<LI>£100 Setup Fee
<LI>£160 Annual Fee
<LI>No Monthly Fee
<LI>Includes ITA Trading Account
</UL>
	<center><a href=https://secure.worldpay.com/app/splash.pl?Pid=62347 target=_blank class=link
onMouseOver="status='http://www.worldpay.com'; return true" onMouseOut="status=''; return true"><B>Apply now</b></a></center>
	</td>
	<td><CENTER><B>Recommended package #2 for UK stores</b></CENTER>
	<UL><LI>3% Discount Rate
<LI>Free Bank Transfers
<LI>No Setup Fee
<LI>No Monthly Fee
<LI>Incl 3rd Party Merch Acct
</UL>
	<center><a href=https://www.moneybookers.com/app/?rid=291696 target=_blank class=link
onMouseOver="status='http://www.moneybookers.com'; return true" onMouseOut="status=''; return true"><B>Apply now</b></a></center>
	</td>
	</tr>
	<tr bgcolor='#FFFFFF'><td colspan=2 align=center><font size=1><a href=ukmerchants.asp class=link><font size=1>Click here to see other UK options</font></a>
<% else %>
	<tr bgcolor='#FFFFFF'><td><CENTER><B>Recommended package #1 international high volume stores</b></CENTER>
	<UL><LI>Rates depend on country store is based in but are generally 3%-4%
</UL>
	<center><a href=https://secure.worldpay.com/app/splash.pl?Pid=62347 target=_blank class=link
onMouseOver="status='http://www.worldpay.com'; return true" onMouseOut="status=''; return true"><B>Apply now</b></a></center>
	</td>
	<td><CENTER><B>Recommended package #2 for international low volume stores</b></CENTER>
	<UL><LI>$49 Application Fee
<LI>No Monthly Fee
<LI>5.5% Discount Rate
<LI>45¢ Transaction Fee
</UL>
	<center><a href=http://www.2checkout.com/cgi-bin/aff.2c?affid=74224 target=_blank class=link
onMouseOver="status='http://www.2checkout.com'; return true" onMouseOut="status=''; return true"><B>Apply now</b></a></center>
	</td>
	</tr>
	<tr bgcolor='#FFFFFF'><td colspan=2 align=center><font size=1>
	<% end if %>

 <tr bgcolor='#FFFFFF'><TD class=instructions>Please note that 3rd party rates are subject to change, these are the rates most recently quoted to our staff but we cannot be responsible for errors.  Non US accounts are highly susceptible to rate changes please contact the provider for current rates.</td></tr>



<%
createFoot thisRedirect, 0


%>


