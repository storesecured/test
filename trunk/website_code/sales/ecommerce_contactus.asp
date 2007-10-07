<%
title = Name&" Contact Form"
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
        sSubject=Name&" contact form from " & fromname
        Send_Mail fromemail,sSales_email,sSubject,message
        response.write "You will be contacted shortly."
    else
        response.write "You did not enter a valid email address.  Please use your browsers back button and try again."
    end if
else
%>
<table width=100%><tr><td valign=top>
StoreSecured Inc<BR>
<%=Name %><br>
10272 Aviary Drive<br>
San Diego, CA 92131<br>
United States<BR></td><td align=right>
<!-- BEGIN LivePerson Button Code --><a href='https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&byhref=1&imageUrl=https://manage.storesecured.com/images/liveperson' target='chat7400929'  onClick="javascript:window.open('https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&imageUrl=https://manage.storesecured.com/images/liveperson&referrer='+escape(document.location),'chat7400929','width=472,height=320');return false;" ><img src='https://server.iad.liveperson.net/hc/7400929/?cmd=repstate&site=7400929&&ver=1&imageUrl=https://manage.storesecured.com/images/liveperson' name='hcIcon' width=165 height=60 border=0></a></a><!-- END LivePerson Button code -->

</td></tr></table>
<br>
Current customers please <a class=link href=http://<%=manage_server %>>login</a>
 to the admin interface to submit a support request.
<BR><BR>
<table border=1 cellspacing=0 cellpadding=5>
<tr><td><b>Email</b></td><td><%=sSales_email %></td><td>Anytime</td></tr>
<tr><td><b>Toll Free #</b></td><td>866-324-2764 (Continental US)</td><td rowspan=2>8:30 - 14:30 PST Mon-Fri<BR>Phones Closed Sat-Sun</td></tr>
<tr><td><b>Alternate #</b></td><td>619-501-0585 (International)</td></tr>
<tr><td><b>Fax #</b></td><td>888-750-6936</td><td>Anytime</td></tr>
<tr><td><b>Live Chat</b></td><td><a href=livechat.asp class=link>Click here</a></td><td>4:30 - 20:00 PST Mon-Fri<BR>8:30 - 17:00 PST Sat-Sun</td></tr>
</table>
<BR><BR>
<B>Or use the form below:</b><BR>
<BR>
<form method="POST" action="ecommerce_contactus.asp" name="ContactUs">
        <table>
        <tr><td>Name</td>
        <td><input type=text name=fromname size=60></td></tr>
        <tr><td>Email</td>
        <td><input type=text name=fromemail size=60></td></tr>
        <tr><td colspan=2><textarea rows="5" name="message" cols="50"></textarea></td></tr>
        <tr><td colspan=2><input type="submit" name=Update value="Send"></td></tr>
        </table>
</form>
Before contacting us check and see if your question is answered in our <a href=http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/front_page!PAGETYPE>FAQ</a>
<BR><BR>
If you have forgotten your password please <a href=http://<%=manage_server %>/forgot_password.asp>click here</a>
<BR><BR>
If you would like to exchange links please <a href=http://<%=sDefaultUrl %>/links/add_link.html>click here</a>

<BR>
<% end if %>

<!--#include file="footer.asp"-->
