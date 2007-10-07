<!--#include virtual="common/connection.asp"-->
<!--#include file="pagedesign.asp"-->
<%

sFormAction = "http://"&manage_server&"/Login_Store_Action.asp"
sTitle = "Login"
thisRedirect = "login_store.asp"
sMenu = "none"
sTopic = "Login"
sFocus = "Store_Id"
createHead thisRedirect  %>
		 <TR bgcolor=FFFFFF>

		 <td class="inputname"><B>Site #</b></td>
			 <td class="inputvalue">
			 <input type="text" name="Store_Id" size="60" value="<%= request.cookies("STORESECURED")("Store_Id") %>">
			 <input type="hidden" name="Store_Id_C" value="Re|Integer|||||Site #" size="20">
			 <% small_help "Store_Id" %></td>
		 </tr>
                 <TR bgcolor=FFFFFF>

		 <td class="inputname"><B>Username</b></td>
			 <td class="inputvalue">
			 <input type="text" name="Store_User_Id" size="60" value="<%= request.cookies("STORESECURED")("Store_User_Id") %>">
			 <input type="hidden" name="Store_User_Id_C" value="Re|String|||||Username" size="20">
			 <% small_help "Login" %></td>
		 </tr>

	 <TR bgcolor=FFFFFF>

			 <td class="inputname"><B>Password</b</td>
			 <td class="inputvalue">
			 <input type="password" name="Store_Password" size="60" value="<%= request.cookies("STORESECURED")("Store_Password") %>">
			 <input type="hidden" name="Store_Password_C" value="Re|String|||||Password" size="20">
			 <% small_help "Password" %></td>
		 </tr>
          <TR bgcolor=FFFFFF>

			 <td class="inputname"><input type="checkbox" name="Save_Info" value="TRUE" <%= request.cookies("STORESECURED")("Save_Info") %>>
				</td>
			 <td class="inputvalue">
                          <B>Save my login information for future visits</b>
			 <% small_help "Save Info" %></td>
		 </tr>

	 <TR bgcolor=FFFFFF>

			 <td width="17%" align="center">&nbsp;</td>
			 <td width="84%" align="left" valign="top"><br>
			 <input type="hidden" name="ReturnTo" value="<%= fn_get_querystring("ReturnTo") %>">
			 <input type="submit" class="Buttons" value="Login" name="I1"><input type="submit" class="Buttons" value="Secure" name="Secure"></td>
			 <td width="84%" align="left" valign="top">&nbsp;</td>
		 </tr>

		 <TR bgcolor=FFFFFF><TD height=10 colspan=3><BR><a href=forgot_password.asp class=link>Forgot your Site #, Username or Password</a></TD></TR>
		 

     <% if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"login_store")>0 or instr(lcase(Request.ServerVariables("HTTP_REFERER")),"www.easystorecreator.com")>0 then %>
		 <TR bgcolor=FFFFFF><TD colspan=3><BR><BR>

		 <table><TR bgcolor=FFFFFF><td bgcolor=red align=center><B><font color=white>Your browser has disabled session cookies.  This site uses temporary, session-based cookies that are created when you log in and automatically erased when you close your browser.</font></b></td></tr></table>




<BR><BR>Please enable session cookies in your browser and then login.
<BR><BR>



To enable session cookies, select your browser from the following list and follow the accompanying instructions.	 If you browser is not listed, then please consult the help menu in the browser for assistance.


<BR><BR><B>Internet Explorer 6</b>
<UL><LI>Select the "Tools" option from your browser’s menu and then choose "Internet Options" (opens a new window).

<LI>Choose the "Privacy" tab and then adjust the settings to "Medium High" or lower.

<LI>Click the "Apply" button.

<LI>If your browser is still refusing to accept session cookies:

<LI>Select the “Tools” option from your browser’s menu and then choose “Internet Options” (opens a new window).

<LI>Go back into the "Privacy" tab and click the "Advanced" button (opens a new window).

<LI>Choose the "Override automatic cookie handling" option and then choose "Always allow session cookies" option.

<LI>Click "OK" followed by "Apply" to save your new settings.
</uL>

<BR><BR><B>Internet Explorer 5</b>
<UL><LI> 		 Select the "Tools" option from your browser’s menu and then choose "Internet Options" (opens a new window).

<LI>Choose the "Security" tab and then click on the "Internet" web content zone.

<LI>Adjust the security level to "Medium" and click "Apply", or click on "Custom Level" and scroll down to "Cookies" and ensure that "Allow per-session cookies (not stored)" is Enabled

<LI>Click "Okay" followed by "Apply".
</uL>

<BR><BR><B>Internet Explorer 4</b>
<UL><LI> 		  Select the "View" option from your browser’s menu and then choose "Internet Options" (opens a new window).

<LI>Choose the "Advanced" tab and then scroll down to the section on "Security".

<LI>Under Security are options for "Cookies"; select "Always accept cookies".

<LI>Click "Apply" in the bottom right of the current window, followed by "Okay".
</uL>

<BR><BR><B>Internet Explorer 5 for Macintosh</b>
<UL><LI> 		  Select the "Edit" option from your browser’s menu and then choose "Preferences" (opens a new window).

<LI>Choose the arrow icon beside the "Receiving Files" option from the left-hand scrolling list, such that it points downwards - displaying a further list of options.

<LI>Click on the "Cookies" option.

<LI>Check the "When receiving cookies" option on the right-hand side of the window and ensure that it is NOT set to "Never Accept".

<LI>Click "OK" to save your new settings.
</uL>

<BR><BR><B>Netscape 6+</b>
<UL><LI> 		  Select the "Edit" option from your browser's menu and then choose "Preferences" (opens a new window)

<LI>Choose "Privacy & Security" from the left-hand category list (expands the menu to display additional options.

<LI>Select "Cookies" from these new options. In the right-hand side of this window you have a choice, choose either Enable all cookies or Enable cookies for the originating web site only.
</uL>

<BR><BR><B>Netscape 4 or 5</b>
<UL><LI> 		  Select the "Edit" option from your browser's menu and then choose "Preferences" (opens a new window).

<LI>Choose "Advanced" from the left-hand category list.

<LI>Choose either Accept all cookies or Accept only cookies that get sent to the originating web site in the Cookies box on the right side of the window.
</uL>


		 </TD></TR>
		 <% end if %>
		 <TR bgcolor=FFFFFF>
<% createFoot thisRedirect,0 %>
