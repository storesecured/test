<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Google Checkout Setup"
thisRedirect = "google_setup.asp"
Secure = "https://"&Secure_name&"/"

sMenu="general"
createHead thisRedirect  %>

<TR bgcolor='#FFFFFF'><td>
<B>Google Checkout Setup</b><BR>
1. Signup for a Google Checkout account using the link below
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><a href="http://checkout.google.com/sell?promo=sestoresecured" target=_blank class=link>Google Checkout Signup</a></I>
<BR><BR>
2. From inside of your Google Checkout account go to
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Integration-->Settings</I>
<BR>
3. Check the the box labelled
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>For extra security, my company will only post digitally signed XML shopping carts</I>
<BR>
4. Enter the following url in the box labelled API callback URL
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/google/callback2.asp</I>
<BR>
5. Choose the type of the return as
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>XML</i>
<BR>
6. Hit Save to save your changes
<BR>
7. From inside the StoreSecured interface go to 
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>General-->Payments-->Gateway</I>
<BR>
8. Check the box labelled
<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Enable Google Checkout</I>
<BR>
9. Enter your Google Merchant Id, Google Merchant Key and currency in the boxes provided
<BR>
10. Hit Save
<BR><BR>
A Google Checkout button will now be shown just below the shopping cart view.  
Google Checkout is not a regular payment method and will not be part of the normal 
checkout process.
<BR><BR>
<B>Limitations of Google Checkout</B>
<UL><LI>Items cannot be shipped to or from multiple locations
<LI>Realtime shipping calculations are not supported
<LI>Duplicate calculated shipping methods, ie those with the same name as other methods will have their id number appended to the name.
<LI>If there are no shipping methods available Google will assign default shipping and taxes
<LI>Promotions are not supported
<LI>Rewards are not supported
<LI>Special customer group pricing is not supported
</UL>

<B>Frequently Asked Questions</b>
<BR><BR><i>Why are there so many limitations on Google Checkout?
</i><BR>Google Checkout is unlike any other payment gateway.
Google requires that the user be transferred directly from the shopping cart 
to the Google Checkout and the remainder of the checkout process is completed
entirely from the Google Checkout webpages.  The user is not passed back to 
the store until the purchase is fully completed.  This presents a few problems 
because there are incompatibilities between what logic Google Checkout supports 
and what logic the store builder supports.  Functionality that the store
supports but Google Checkout does not is therefore lost.
<BR><BR><i>What are the advantageous of using Google Checkout?</b>
</i><BR>Google is offering free processing until the end of this year.  
In addition they charge no monthly fees and once fees are charged they 
will credit a percentage of those fees against your Google Adwords account.
In addition like PayPal payments Google Checkout keeps the users payment 
information private which appeals to many customers who do not wish to 
share that information with all merchants.
<BR><BR><i>If I enable Google Checkout how does it work?
</i><BR>When Google Checkout is enabled a small Google Checkout graphic 
is added to the bottom of your shopping cart view page.  This graphic 
is automatically inserted and is provided by Google.  When the graphic 
is clicked the user will be taken to the Google Checkout page to 
complete the purchase.  For users familiar with Paypal Express Checkout 
it is similar except that the user does not return back to the store after 
they login at the 3rd party site.
<BR><BR><i>How can I apply for a Google Checkout account?
</i><BR>You may apply by clicking on the Google Checkout signup link 
below, please note that you must wait until you are approved to be a part of the Beta
Group or the functionality is offered to all stores before you can
begin using the account even if you apply today.

</td></tr>


<% createFoot thisRedirect, 0 %>


