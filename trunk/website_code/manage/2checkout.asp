<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "2Checkout Instructions"
thisRedirect = "2checkout.asp"
Secure = "https://"&Secure_Name&"/"
sMenu="general"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up 2Checkout Version 1</td></tr>
		<TR bgcolor='#FFFFFF'><TD nowrap>1. Login to your 2Checkout account<BR>
		2. Select Shopping Cart --> Cart Details<BR>
		3. Set Return to a routine on your site after credit card processed to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>Yes</I><BR>
		4. Set Return URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/2checkout/2checkout.asp?Store_Id=<%= Store_ID %></I><BR>
		5. Select save changes
		<BR><BR>
		<B>Instructions for setting up 2Checkout Version 2</b>
                <BR>1. Login to your 2Checkout account<BR>
		2. Select Look and Feel<BR>
		3. Set Direct Return to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I>No</I><BR>
		4. Set Approved URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/2checkout/2checkout.asp?Store_Id=<%= Store_ID %></I><BR>
		5. Set Pending URL to:<BR>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<I><%= Secure %>include/2checkout/2checkout.asp?Store_Id=<%= Store_ID %></I><BR>
		6. Select save changes
		<BR><BR>
      <a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
      </td></tr>


<% createFoot thisRedirect, 0 %>


