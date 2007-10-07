<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Paypal Pro Setup"
thisRedirect = "paypal_pro.asp"
Secure = "https://"&Secure_Name&"/"
sMenu="general"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up Paypal Pro</td></tr>
		<TR bgcolor='#FFFFFF'><TD>
Please follow the instructions to use PayPal Pro
<BR>1) Visit General->Payment-->Configure Gateway
<BR>2) Select Use PayPal Website Payment Pro
<BR>3) Enter the Paypal API User name
<BR>4) Enter the Paypal password
<BR>5) Enter the 3 digit ISO currency code, <a href=http://www.xe.com/iso4217.php class=link target=_blank>Click here for a list of currency codes</a>
<BR>6) Enter the Paypal Security Key
<BR>7) Hit Save
<BR><BR>
<BR>Note: You must have an Activated Pro-Level account to setup.
<BR>
<BR>How do I get PayPal the Paypal Information?
<BR>1) Log in to your PayPal Premier or Business account.

<BR>2) Click the Profile subtab located in the top navigation area.

<BR>3) Click the API Access link under the Account Information header.

<BR>4) Click the Request API Credentials link.

<BR>5) Choose the Radio button for API Signature and agree to the terms of use.

<BR>6) Write down the values for API Username, Password and Signature or print the page.

<BR>7) Copy the information you have written down into the Storesecured site where indicated.
</td></tr>


<% createFoot thisRedirect, 0 %>


