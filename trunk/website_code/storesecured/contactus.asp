<%
title = "Easystorecreator Contact Form"
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

    if instr(fromemail,"@") > 0 and instr(fromemail," ") = 0 then
       Set Mail = Server.CreateObject("Persits.MailSender")
       Mail.From = fromemail
       Mail.AddAddress sSales_email
       Mail.Subject = Name&" contact form from " & fromname
       Mail.Body = message
       Mail.Queue = True
       Mail.Send
       response.write "You will be contacted shortly."
    else
       response.write "You did not enter a valid email address.  Please use your browsers back button and try again."
    end if
else
%>

StoreSecured.com<br>
10272 Aviary Drive<br>
San Diego, CA 92131<br>
United States<BR>
<br>
<B>Call toll free:</b> 866-324-2764 (Continental US Callers Only)<br>
<B>Alternate number:</b> 619-501-0585 (Local and International Calls)<br>
<B>Fax:</b> 888-750-6936<br>
<B>Or use the form below:</b><BR>
<BR>
<B>Phone support hours:</b> 8:30AM - 2:30PM PST Mon-Fri
<form method="POST" action="Contactus.asp" name="ContactUs">
        <table>
        <tr><td>Name</td>
        <td><input type=text name=fromname size=20></td></tr>
        <tr><td>Email</td>
        <td><input type=text name=fromemail size=20></td></tr>
        <tr><td>To</td>
        <td><%= sSales_email %></td></tr>
        
        <tr><td colspan=2><textarea rows="5" name="message" cols="50"></textarea></td></tr>
        <tr><td colspan=2><input type="submit" name=Update value="Send"></td></tr>
        </table>
</form>
<BR>
Click here to have a <a href=livechat.asp>live chat</a> with an operator
<BR><BR>
<% end if %>

<!--#include file="footer.asp"-->
