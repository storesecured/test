<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "USPS Instructions"
thisRedirect = "usps.asp"
sMenu="general"
createHead thisRedirect  %>



		<TR><td><B>Instructions for getting USPS setup with Easystorecreator</td></tr>
		<TR><TD nowrap>1. Go to <a href=http://www.uspsprioritymail.com/et_regcert.html class=link target=_blank>USPS registration</a><BR>
		2. Enter company information to complete registration<BR>
		3. Login to UPS.com using the login you just created<BR>
		4. Choose the get tool link under UPS Rates & Service Selection<BR>
		5. Enter company information to get a developers key<BR>
		6. Choose get access key from left menu<BR>
		7. Choose xml license key<BR>
		8. Enter the developers key that was mailed to you<BR>
		9. Enter company information to get your access key<BR>
		10. Login to Easystorecreator go to General Settings-->Shipping<BR>
		11. Choose realtime shipping<BR>
		12. Enter you UPS username, password and access license<BR>
		13. Check the box marked "Use UPS"<BR>
		14. Select Save

		</td></tr>

	  <TR><TD></td></tr>

<% createFoot thisRedirect, 0 %>


