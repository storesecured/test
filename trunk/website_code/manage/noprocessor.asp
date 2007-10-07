<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "No Processor"
thisRedirect = "noprocessor.asp"
Secure = Replace(Site_Name,"http://","https://")
sMenu="general"
createHead thisRedirect  %>



		<tr bgcolor="#FFFFFF"><td><B>Instructions for using No Processor</td></tr>
		<tr bgcolor="#FFFFFF"><TD>Selecting no processor means that you will either be processing credit cards offline or you will not be accepting
		credit card payments.  
      <BR><BR><B>IMPORTANT: this does NOT mean that StoreSecured will process payments for you.</b><BR>
      <BR>
		If you do not wish to accept credit card payments in your store go to General Settings-->Payment Methods
		and ensure that you have unchecked the boxes for credit cards, ie Visa, Mastercard, Discover, American Express and Diners Club<BR>
		<BR><BR>If you plan on processing credit cards offline you must have a processor already setup to do so.  In order to view credit card details
		for a particular order you must be logged in securely otherwise credit card numbers are masked to ensure data security.  To login securely 
		select the Secure button during login or change the http:// in the address bar of your browser to https://
      <BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
		</td></tr>

	  <tr bgcolor="#FFFFFF"><TD></td></tr>

<% createFoot thisRedirect, 0 %>


