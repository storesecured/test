<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Refer Friends"
sFullTitle = "My Account > Refer Friends"
thisRedirect = "referral.asp"
sMenu="account"
sQuestion_Path = "import_export/my_account/refer_friends.htm"
createHead thisRedirect  %>


      <TR bgcolor='#FFFFFF'><TD>
      <B>1. Send your personalized referral link in an email</b>
				<br>
				<br>Send a short Referral Email to merchants you know and include the link below, or <a href="mailto:?Subject=Start%20an%20online%20store%20today%20with%20Easystorecreator%21&Body=When%20you%20signup%20with%20Easystorecreator%2C%20you%20can%20create%20your%20own%20online%20store%20instantly.%20Accept%20credit%20cards%20in%20your%20store%20and%20process%20them%20in%20real-time%20.%20%20Easystorecreator%20is%20the%20fastest%20way%20to%20open%20your%20doors%20to%20millions%20of%20internet%20users%20worldwide.%0D%0A%0D%0ABest%20of%20all%2C%20it%27s%20completely%20free%20to%20sign%20up%21%0D%0A%0D%0ATo%20sign%20up%20or%20learn%20more%2C%20click%20here%3A%20%0D%0A%0D%0A%20http://www.easystorecreator.com?Id=<%=Store_Id %>" class=link>click here</a> &#32;to get a sample Referral Email with your personalized referral link included.
					<br><br>
				<INPUT type="hidden" name="cmd" value="_refer-email">
				<INPUT type="text" size="60" value="http://www.easystorecreator.com?Id=<%= Store_Id%>">
				</FORM></li>
<br><br>
				<B>2. Add a referral link to your website</b>
				<br><br>
				Put a referral link on your website by copying the code below and pasting it into your website's HTML code:<br>
				<br>
				<br>
				<textarea NAME="LogoHTML" Rows=4 COLS=60 wrap="hard"><!-- Begin Easystorecreator Refer --><A HREF="http://www.easystorecreator.com?Id=<%= Store_Id%>" target="_blank">Start your own online store today</a><!-- End Easystorecreator Link --></textarea></li>
</ol>
<BR>Each time a new merchant signs up for a service via the link, your account will be automatically credited for one month free Service.
</td></tr>

<% createFoot thisRedirect, 0 %>


