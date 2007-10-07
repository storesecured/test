<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "PayPal Pro Important Updates"
thisRedirect = "paypal_christmas.asp"
sMenu="general"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td>
<table><tr><td width='70%'><B>Special PayPal Express Checkout Offer for StoreSecured Merchants</b><BR>
Free transaction processing on Express Checkout transactions until 1/31/2008.  Plus receive a $100 credit in a Yahoo! Sponsored Search account.
<a class=link target=_blank href="http://www.paypal-promo.com/searchmarketing/tracking/media_source/easystore.html">
Click here to Signup with Paypal Express Checkout</a></td><td width='30%' align=right><a href="http://www.paypal-promo.com/searchmarketing/tracking/media_source/easystore.html"><img src="http://www.paypal-promo.com/searchmarketing/images/banner/150x100_a.gif" border=0></a>
</td></tr></table>
<B>PayPal Pro Important Update 4/22/2007</b>
<BR><BR>StoreSecured has recently released a new version of software to support PayPals newest programming interface.
This PayPal update requires all PayPal Pro merchants to update their PayPal settings.  The old settings will cease to be supported 
after 5/31/2007.
<BR><BR><font color=red>Please note that this update DOES NOT apply to PayPal Standard.  If you are not using PayPal Pro or Express Checkout please disregard this notice.</font>
<BR><BR>To update your PayPal Pro settings please follow the directions below:
<OL>
<LI>Log in to your PayPal Premier or Business account.
<LI>Click the Profile subtab located in the top navigation area.
<LI>Click the API Access link under the Account Information header.
<LI>In the Request API Credentials box click View or Remove Credentials
<LI>Choose the Radio button for API Signature and agree to the terms of use.
<LI>Click on the Remove button, PayPal will ask you to confirm by reselecting Remove
<LI>You will be taken back to the previous page.
<LI>Click on the Request API Credentials link
<LI>Select the radio button for API Signature
<LI>Check the box indicating you agree to PayPals Terms of use
<LI>Select the Submit button
<LI>Write down the values for API Username, Password and Signature or print the screen.
<LI>Copy the API Username, Password and Signature into the corresponding boxes in the StoreSecured interface under General-->Payments-->Gateway and hit Save
</OL>
<B>Important Notes</b>
<BR>*We recommend performing at least one test transaction to ensure that you have copied the values correctly.   It is important to enter your values into StoreSecured immediately after creating them in PayPal.  Your old values will cease to work with PayPal once the new values are created.
<BR>**Some merchants have experienced a lag time of approximately 15-30 minutes between when the information was received from PayPal and when PayPal would start accepting that new information as valid.
If it does not work immediately please wait between 15-30 minutes and try again.
</td></tr>


<% createFoot thisRedirect, 0 %>


