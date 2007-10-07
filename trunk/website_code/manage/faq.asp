<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Frequently Asked Questions"
thisRedirect = "faq.asp"
sMenu="none"
addPicker=1
createHead thisRedirect
%>
	 <TR><TD><BR>1. <a class="link" href=#items-missing><B>How come my items don't show up when I browse my store?</b></a><BR></TD></TR>
	 <TR><TD><BR>2. <a class="link" href=#members-only><B>When I try to browse my store it always takes me to the login page?</b></a><BR></TD></TR>
	 <TR><TD><BR>3. <a class="link" href=#flyover-help><B>There used to be help on the side of all of the screens but now its gone?</b></a><BR></TD></TR>
	 <TR><TD><BR>4. <a class="link" href=#small-image><B>Why do my small images show up so large on the item view?</b></a><BR></TD></TR>
	 <TR><TD><BR>5. <a class="link" href=#link-popup><B>Sometimes I click on a link and nothing happens?</b></a><BR></TD></TR>
	 <TR><TD><BR>6. <a class="link" href=#paypal-signup><B>I am using a paypal account to process credit cards, is there
	 some way that the shoppers can enter their credit card information into my store and I can process it so they don't
	 have to create a login?</b></a><BR></TD></TR>
	 <TR><TD><BR>7. <a class="link" href=#verified-orders><B>How do some orders get marked as verified, what does that mean?</b></a><BR></TD></TR>
	 <TR><TD><BR>8. <a class="link" href=#update-stock><B>I have indicated that I want stock updated automatically but after placing an item into
	 my cart the quantity in stock is still the same?</b></a><BR></TD></TR>
	 <TR><TD><BR>9. <a class="link" href=#credit-cards><B>When I view order detail I do not see the credit card number used
	 to make the purchase?</b></a><BR></TD></TR>
	 <TR><TD><BR>10. <a class="link" href=#attribute><B>What is an item attribute and why would I want to use one?</b></a><BR></TD></TR>
	 <TR><TD><BR>11. <a class="link" href=#ssl><B>I noticed that during the checkout process the address bar says something different than my domain
	 name, why does the address change?</b></a><BR></TD></TR>
	 <TR><TD><BR>12. <a class="link" href=#ssl-compatible><B>One of my shoppers have reported that they did not see the padlock
	 at the bottom of their browser during checkout?</b></a><BR></TD></TR>
	 <TR><TD><BR>13. <a class="link" href=#export-text><B>I checked the export orders/customers/items box where is my text file?</b></a><BR></TD></TR>
	 <TR><TD><BR>14. <a class="link" href=#carriage-return><B>I typed in text for my page content with different paragraphs and when it gets displayed it looks like one big paragraph?</b></a><BR></TD></TR>
	 <TR><TD><BR>15. <a class="link" href=#aol-email><B>I am not receiving any of the email notifications when a user registers, an order takes place etc?</b></a><BR></TD></TR>
	 <TR><TD><BR>16. <a class="link" href=#cancel><B>How do I cancel my account?</b></a><BR></TD></TR>
	 <TR><TD><BR>17. <a class="link" href=#other-questions><B>I need help with a question not answered here, can you help me?</b></a><BR></TD></TR>

	 <TR><TD height=100></TD></TR>
	 <TR><TD><BR><B><a name=items-missing>How come my items don't show up when I browse my store?</a></b><BR>
	 If you items aren't showing up after adding you should check to make sure that you assigned
	 them to at least one department from the <a class="link" href="edit_items.asp">Edit Items</a> page.	If they are not assigned to a department they have no place to go.
	 The second thing to check is to make sure that the visible checkbox is checked.  If it is not checked
	 items will be hidden from shoppers.</TD></TR>
	 <TR><TD><BR><B><a name=members-only>When I try to browse my store it always takes me to the login page?</a></b><BR>
	 Your store is probably set up as non-public.  A public store implies that all shoppers can browse your
	 store before registering, while a non-public store implies that shoppers must register before they can even
	 browse your store.	To change this setting go the <a class="link" href="misc_settings.asp">Misc Settings</a> page.</TD></TR>
	 <TR><TD><BR><B><a name=flyover-help>There used to be help on the side of all of the screens but now its gone?</a></b><BR>
	 You can toggle the context sensitive help on and off.  We recommend turning it off once you know your way around
	 since the pages tend to load more quickly this way.	But if you need it back there is a link at the bottom right of each page
	 that says either Show Help or Hide Help (depending on your current settings).  Click the link to toggle the help on or off
	 or click here to <a class="link" href=hide_help.asp>Show Help</a>.</TD></TR>
	 <TR><TD><BR><B><a name=small-image>Why do my small images show up so large on the item view?</a></b><BR>
	 We do not resize your images once they are uploaded.  All images are shown in their original size.	If you would like
	 for your images to be smaller you should resize them before uploading.  There are several good free resizing programs.
	 We recommend using Irfanview.</TD></TR>
	 <TR><TD><BR><B><a name=link-popup>Sometimes I click on a link and nothing happens?</a></b><BR>
	 There are several reasons why a link might not be doing anything.  First you must have a browser which supports
	 javascript.  We suggest using the latest version of Internet Explorer.  Second, many of our secondary windows, ie
	 color picker, image picker, file picker, editor and preview windows rely on opening in a second window.  If you have
	 a popup blocker installed you should consider disabling it so that these windows can popup.</TD></TR>
	 <TR><TD><BR><B><a name=paypal-signup>I am using a paypal account to process credit cards, is there
	 some way that the shoppers can enter their credit card information into my store and I can process it so they don't
	 have to create a login?</a></b><BR>
	 No this is not possible with paypal.	If you don't want your shoppers to have to register you will need to use another
	 processor. There is no way that you can register a shopper becuase they ask many personal questions which only the
	 cardholder would know.  If you would like more information about other credit card processors go to our
	 <a class="link" href="merchants.asp">Payment Processor</a> page</TD></TR>
	 <TR><TD><BR><B><a name=verified-orders>How do some orders get marked as verified, what does that mean?</a></b><BR>
	 A verified order indicates that payment has been verified. If you are processing credit cards using a real-time processor
	 and the payment is accepted it will be instantly verified by our system.	For other types of payment such as fax, check,
	 etc it will be the store admins responsibility to mark orders as verified.</TD></TR>
	 <TR><TD><BR><B><a name=update-stock>I have indicated that I want stock updated automatically but after placing an item into
	 my cart the quantity in stock is still the same?</a></b><BR>
	 In stock quantities are not updated by our system until the order is marked as verified.  If you are worried about selling
	 items you do not have we recommend placing the sell limit higher than 0.</TD></TR>
	 <TR><TD><BR><B><a name=credit-cards>When I view order detail I do not see the credit card number used
	 to make the purchase?</a></b><BR>
	 If you are using a payment processor to process your credit card transactions our system will store the actual credit
	 card information along with authorization and verification numbers for transactions which can be used to lookup the
	 actual transaction on the processors website.	To view sensitive information such as credit card information you must 
	 be logged in securely.  This is done to ensure the security of your cusomters information.</TD></TR>
	 <TR><TD><BR><B><a name=attribute>What is an item attribute and why would I want to use one?</a></b><BR>
	 Attributes are modifiers to an item.	For instance if you sold shirts you could create 4 shirt inventory items
	 or if they were all the same shirt in different colors you could add an attribute color to allow the shopper to
	 select from when they added the shirt to their cart.  You might also want to offer them the attribute size.  In fact
	 you can even charge different prices for different attributes.  For instance many shops charge extra for extra large
	 sizes or hard to find items.  Most stores find it is easier to keep a lesser number of items with attributes.</TD></TR>
	 <TR><TD><BR><B><a name=ssl>I noticed that during the checkout process the address bar says something different than my domain
	 name, why does the address change</a></b><BR>
	 The address bar changes so that the shoppers checkout can be done via a secure connection.	We as a company only
	 have the right to secure domains which end in easystorecreator.com.  Thus during checkout we will redirect the shopper
	 to one of our subdomains.  Your store will look exactly the same.  The only difference will be that it is secured and
	 the store address is slightly different.  If you want your visitors to stay on your domain name the whole time you must
	 purchase a secure site license and have us install it on our server.  The price for this service is $50 per year.  We feel
	 this is not needed as 99% of shoppers will never even notice a difference.</TD></TR>
	 <TR><TD><BR><B><a name=ssl-compatible>One of my shoppers have reported that they did not see the padlock
	 at the bottom of their browser during checkout?</a></b><BR>
	 You should assure all customers that regardless of whether they see the lock or not they are protected as long
	 as the address bar starts with https instead of http.  While our ssl license is recognizable by 99% of browsers
	 there is 1% of the time where it is not recognized.	If shoppers are wary of purchasing we recommend you advise them
	 to upgrade their browser to something more recent.	It should be noted that no ssl certificate is 100% recognizable
	 by all browsers.  </TD></TR>
	 <TR><TD><BR><B><a name=export-text>I checked the export orders/customers/items box where is my text file?</a></b><BR>
	 All exported files are always available in the <a class="link" href=export_items.asp>export items</a> page.  If you
	 don't see your file make sure that after checking the box you performed the search using the provided button or try
	 refreshing the export page.</TD></TR>
	 <TR><TD><BR><B><a name=carriage-return>I typed in text for my page content with different paragraphs and when it gets
	 displayed it looks like one big paragraph?</a></b><BR>
	 In HTML special codes are needed to perform a line break or text formatting.  If you do not know these codes and you simply
	 hit the enter key to try and format the text it will not be displayed correctly.  If you don't know html we highly recommend
	 you use the editor link whenever possible.	The editor link will popup a new window which will allow you to format the text
	 as if you were in microsoft word.	When using the editor it will translate your intentions into HTML code for you.</TD></TR>
	 <TR><TD><BR><B><a name=aol-email>I am not receiving any of the email notifications when a user registers, an order takes place etc?</a></b><BR>
	 If this is happening to you its probably because you are using an aol email address. AOL will not accept email addressed
	 to the same person it is being sent to.	All of your email is sent from the email you have listed as your contact email address
	 so if you use the same email address for your notifications it will not work.  We suggest using 2 different emails, one as your
	 store email and one as the notification email, or use another email service besides AOL.</TD></TR>
	 <TR><TD><BR><B><a name=cancel>How do I cancel my account?</a></b><BR>
	 You may <a href=cancel_store.asp class=link>cancel your account</a> at any time.  Please note that by requesting your store to be cancelled Easystorecreator
	 will be completely removing your store from our servers.  Once a store is removed it cannot be reactivated it must be recreated.</TD></TR>

	 <TR><TD><BR><B><a name=other-questions>I need help with a question not answered here, can you help me?</a></b><BR>
	 Sure, for the fastest response please fill out a <a href=support.asp class=link>support request</a>.  Make sure to mention your store id:<B><%=Store_Id %></b> in all
	 correspondence for the fastest possible response.</TD></TR>

	 <TR><TD height=600></TD></TR>
<% createFoot thisRedirect, 0%>
