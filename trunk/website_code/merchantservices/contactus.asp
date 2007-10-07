<%
title = "EasyMerchantServices Contact Form"
description = "Merchant account provider offers merchant account processor functions. Integrates with many merchant account service provider companies. Apply for your small business merchant account today. We'll completely setup your online store for you."
keyword1="merchant account provider"
keyword2="merchant account processor"
keyword3="merchant account service provider"
keyword4="small business merchant account"
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="contactus"

%>
<!--#include file="header.asp"-->
<p class="bodytext">

<%
if request.form("Update") <> "" then
    fromname = Replace(Request.Form("fromname"), "'", "''")
    fromemail = Replace(Request.Form("fromemail"), "'", "''")
    message = Replace(Request.Form("message"), "'", "''")

    Set Mail = Server.CreateObject("Persits.MailSender")
    Mail.From = fromemail
    Mail.AddAddress "sales@easymerchantservices.com"
    Mail.Subject = "EasyMerchantServices contact form from " & fromname
    Mail.Body = message
    Mail.Queue = True
    Mail.Send
    response.write "You will be contacted shortly."
else
%>

EasyMerchantServices.com<br>
10272 Aviary Drive<br>
San Diego, CA 92131<br>
United States<BR>
<br>
<B>Call toll free:</b> 866-324-2764<br>
<B>Email:</b> sales@easymerchantservices.com<BR>
<B>Office Hours:</b> 8:30AM - 2:30PM PST<BR>
<B>Or use the form below:</b><BR>
<form method="POST" action="Contactus.asp" name="ContactUs">
        <table>
        <tr><td>Name</td>
        <td><input type=text name=fromname size=20></td></tr>
        <tr><td>Email</td>
        <td><input type=text name=fromemail size=20></td></tr>
        <tr><td>To</td>
        <td>sales@easymerchantservices.com</td></tr>
        
        <tr><td colspan=2><textarea rows="5" name="message" cols="50"></textarea></td></tr>
        <tr><td colspan=2><input type="submit" name=Update value="Send"></td></tr>
        </table>
</form>

<% end if %>
<!-- BEGIN LivePerson Button Code --><table border='0' cellspacing='2' cellpadding='2'><tr><td align="center"></td><td align='center'><a href='http://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&byhref=1' target='chat7400929'  onClick="javascript:window.open('http://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&referrer='+escape(document.location),'chat7400929','width=472,height=320');return false;" ><img src='http://server.iad.liveperson.net/hc/7400929/?cmd=repstate&site=7400929&&ver=1&category=en;woman;1' name='hcIcon' width=150 height=50 border=0></a></td></tr><tr><td>&nbsp;</td><td align='center'><a href='http://www.liveperson.com' style='text-decoration:none'><font face='Arial' size='-2' color='#333333'><span style='font-size: 10px; font-family: Arial, Helvetica, sans-serif'>Live chat by<font color='#475780'> Live</font><font color='#56A145'>Person</font></span></font></a></td></tr></table><!-- END LivePerson Button code -->
	  

<!--#include file="footer.asp"-->
