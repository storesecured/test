<!--#include file="include/header.asp"-->
<%
redirect = request.form("redirect")
if redirect <> "" then
   fn_redirect redirect
end if

response.write "<form name='Show_big_cart' action='"&Switch_Name&"Cart_Action.asp?Return_To="&server.urlencode(unencode(fn_get_querystring("Return_To")))&"' method=post>"
%>

<TABLE>
	<TR>
		<TD>
		<input type=hidden name="Return_Back" value="0">
      <%= fn_create_action_button ("Button_image_ContinueShopping", "Continue_Shopping", "Continue Shopping") %>
      </TD>
		<TD>&nbsp;&nbsp;</TD>
		<TD>
      <%= fn_create_action_button ("Button_image_Checkout", "Check_Out", "Check Out") %></TD>
	</TR>
</TABLE>
</form>

<!--#include file="include/footer.asp"-->
