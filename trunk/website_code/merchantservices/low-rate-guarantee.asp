<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="low-rate-guarantee.asp"

%>
<!--#include file="header.asp"-->
<% if request.form("Update") <> "" then
    url = Replace(Request.Form("url"), "'", "''")
    name = Replace(Request.Form("name"), "'", "''")
    email = Replace(Request.Form("email"), "'", "''")
    monthly = Replace(Request.Form("monthly"), "'", "''")
    avg = Replace(Request.Form("avg"), "'", "''")
    message = "url="&url&vbcrlf&"monthly="&monthly&vbcrlf&"avg="&avg

    Set Mail = Server.CreateObject("Persits.MailSender")
    Mail.From = email
    Mail.AddAddress "sales@easymerchantservices.com"
    Mail.Subject = "EasyMerchantServices lower rate from " & name
    Mail.Body = message
    Mail.Queue = True
    Mail.Send
    response.write "You will be contacted shortly by a member of our staff."
else %>
<H1>Low Rate Guarantee</h1>
If you find a lower rate for internet processing ANYWHERE on the net we will beat it by 5%.
<BR><BR>
Please fill out the form below to alert us of a lower rate so we can beat it.
<form method=post action=low-rate-guarantee.asp>
<table>
<tr><td>Website address of lower rate</td><td><input type=text name=url></td></tr>
<tr><td>Name</td><td><input type=text name=name></td></tr>
<tr><td>Phone</td><td><input type=text name=phone></td></tr>
<tr><td>Email</td><td><input type=text name=email></td></tr>
<tr><td>Monthly sales</td><td>$<input type=text name=monthly></td></tr>
<tr><td>Average transaction</td><td>$<input type=text name=avg_tran></td></tr>


<tr><td colspan=2><input type=submit name=Update value=Send></td></tr>
</table>
</form>
Terms of guarantee:
<UL><LI>
Advertised rate must include both a merchant account and internet gateway together.
<LI>Advertised rate must have equivalent
terms, ie no cancellation fee, no monthly minimum, no application fee, $50 or less setup fee, etc.
<LI>Advertised rate must be
for an actual merchant account and not a 3rd party or person to person payment system.
</UL>
<% end if %>
<!--#include file="footer.asp"-->
