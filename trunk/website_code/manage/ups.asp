<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "UPS Instructions"
thisRedirect = "ups.asp"
sMenu="shipping"
createHead thisRedirect  %>



		<TR bgcolor="#FFFFFF"><td><B>Instructions for UPS setup</td></tr>
		<TR bgcolor="#FFFFFF"><TD nowrap>
		If you are using UPS for realtime shipping rates only
		<BR>
                1. Enter in as the username wcxtest
                <BR>2. Enter in as the password wcxtest
                <BR>3. Enter in for the access license 0B6818030438401C
                <BR><BR>
                If you are using UPS for shipping labels
			 <BR>1. Go to <a href=https://www.ups.com/servlet/registration?loc=en_US_EC&returnto=http://www.ec.ups.com/ecommerce/techdocs/online_tools.html class=link target=_blank>UPS registration</a><BR>
		2. Enter company information to complete registration<BR>
		3. Login to UPS.com using the login you just created<BR>
		4. Choose the get tool link under UPS Rates & Service Selection<BR>
		5. Enter company information to get a developers key<BR>
		6. Choose get access key from left menu<BR>
		7. Choose xml license key<BR>
		8. Enter the developers key that was mailed to you<BR>
		9. Enter company information to get your access key<BR>
		11. Enter you UPS username, password and access license you have received from UPS<BR>


                 <BR><BR>
                <a href=shipping_class_realtime.asp>Return to shipping setup</a>

		</td></tr>


                <!--
	  	<TR>
			<td>
				<BR><BR><B>Instructions for setting up your UPS Account Number in your UPS profile (required for printing Shipping Labels)
			</td>
		</TR>

		<TR>
			<TD nowrap>
				1. Using the UserId you plan to use for your shipments, log into UPS.com at <a href=http://www.ups.com class=link target=_blank>www.ups.com</a> <BR>
				2. Click on MyUPS located at the top right hand corner of the screen.<BR>
				3. Click on Manage My UPS in the left side-bar navigation.<BR>
				4. Add the Account Number to your ID from the Account Summary page.<BR>
			</TD>
		</TR>
		-->

	  


<% createFoot thisRedirect, 0 %>


