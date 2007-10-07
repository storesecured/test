<!--#include file="include/header.asp"-->

<% 

if redirect <> "" then
    redirect = request.form("redirect")
    fn_redirect "redirect"
end if

sReturn_To = unencode(fn_get_querystring("Return_To"))
if sReturn_To="" then
    sReturn_To = Switch_Name&"items/list.htm"
end if
response.write "<form name='Show_big_cart' action='"&Switch_Name&"cart_action.asp?Return_To="&server.urlencode(sReturn_To)&"' method=post>"
%>

<TABLE>
	<TR>
		<TD>
      <%= fn_create_action_button ("Button_image_ContinueShopping", "Continue_Shopping", "Continue Shopping") %>
      </TD>
		<TD>&nbsp;&nbsp;</TD>
		<TD>
      <%= fn_create_action_button ("Button_image_Checkout", "Check_Out", "Check Out") %></TD>
	</TR>
</TABLE>
</form>

<!--#include file="include/footer.asp"-->
