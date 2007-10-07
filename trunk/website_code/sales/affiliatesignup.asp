<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="affiliatesignup"

%>
<!--#include file="header.asp"-->
<table border="0" width=525 align=left bordercolor="#000066">

			<tr>
					<td colspan="" valign="top" align=center>
							<b>Affiliate Program</b><br>

							</td>
				
				</tr>

			<form method="POST" action="http://refer.EasyStoreCreator.com/create_affiliate_action.asp">
			<tr><td colspan=2 align=left>New Affiliate Signup</td></tr>
			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">First Name</font></td>
					<td width="98%" height="1">
					<input type="text" name="First_name" size="20">
					<input type="hidden" name="First_name_C" value="Re|String|0|50|||First Name"></td>
				</tr>
			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Last Name</font></td>
					<td width="98%" height="1">
					<input type="text" name="Last_name" size="20">
					<input type="hidden" name="Last_name_C" value="Re|String|0|50|||Last Name"></td>
				</tr>

			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Login</font></td>
					<td width="98%">
					<input type="text" name="Login" size="20">
					 <input type="hidden" name="Login_C" value="Re|String|0|20|||Login"></td>
				</tr>

			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Password</font></td>
					<td width="98%">
					<input type="password" name="Password" size="20">
					<input type="hidden" name="Password_C" value="Re|String|0|20|||Password"></td>
				</tr>

			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Password Confirmation</font></td>
					<td width="98%">
					<input type="password" name="password_confirm" size="20">
					<input type="hidden" name="Password_confirm_C" value="Re|String|0|20|||Password Confirmation"></td>
				</tr>
			 <tr>

					<td width="1%" nowrap><font face="Arial" size="2">Email</font></td>
					<td width="98%">
					<input type="text" name="Email" size="20">
					<input type="hidden" name="Email_C" value="Re|Email|0|50|||Email">
					<input type="hidden" name="Affiliate" value="<%= Request.Cookies("EASYSTORECREATOR")("Affiliate") %>">
</td>
				</tr>
				<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Website</font></td>
					<td width="98%">
					<input type="text" name="Website" size="20">
					<input type="hidden" name="Website_C" value="Re|String|0|50|||Email">

				</tr>

			<tr>

					<td colspan=2 align=left><br>
						<input type="submit" border="0" value="Create"></td>
				</tr>
						
			</form>

		</table>
<!--#include file="footer.asp"-->
