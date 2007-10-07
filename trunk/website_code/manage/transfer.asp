<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Transfer Instructions"
thisRedirect = "transfer.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="general"
createHead thisRedirect  %>



		<tr bgcolor='#FFFFFF'><td><B>Instructions for completing transfer</td></tr>
		<tr bgcolor='#FFFFFF'><TD>
		To complete the transfer process you MUST update your domain nameservers.<BR>
                If you do not do the following your domain name will never point at your store.<BR>
                <HR>
                1. Go to the website of the company you originally registered the domain name through (Some of the more popular registrars who you might of used are networksolutions.com, godaddy.com and register.com)<BR>
                2. Login to your account with that company<BR>
                3. Find the option to update your nameservers for the domain name<BR>
                4. Update the nameservers to read <i>NS.RACKSPACE.COM</i> and <i>NS2.RACKSPACE.COM</i><BR>
                5. Wait 24-48 hours after which time your domain name should point at your store<BR>
                6. Once the name points at your store go to Advanced-->domain name in your store manager and click on the link to make your domain name the default.<BR>
                <BR><BR>
                If you are having trouble finding the exact option to update the nameservers with your registrar please contact the registrar directly for further help.  Since each registrars process is slightly different our staff is unable to provide specific guidance for all of the different registrars.


		</td></tr>



<% createFoot thisRedirect, 0 %>


