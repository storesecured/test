<!--#include file="include/header.asp"-->
<!--#include file="include/cart_display.asp"-->

<%
on error resume next
if Redirect_To_Cart <> 0 then
	 checked1="checked"
	 checked0=""
	 sRedirect=1
else
	 checked0="checked"
	 checked1=""
	 sRedirect=0
end if

if fn_get_querystring("Return_To") <> "" then
    Return_To = fn_get_querystring("Return_To")
    'fLocal=instr(Return_To,Switch_Name)
    'if Return_To <> "http:" and (fLocal=0 or (fLocal>0 and instr(Return_To,"items/")>0)) then
    'else
    '    Return_To = Switch_Name&"items/list.htm"
    'end if
elseif request.form("Return_To")<>"" then
    Return_To = request.form("Return_To")
elseif request.ServerVariables("HTTP_REFERER")<>"" then
    Return_To = request.ServerVariables("HTTP_REFERER")&"#"&request.form("Item_Id")
else
    Return_To = Switch_Name&"items/list.htm"
end if

Return_ToLeft=left(Return_To,4)
if Return_ToLeft<>"http" then
	Return_To=""
end if

fn_print_debug "return_to="&return_to

sTotal_cart=fn_get_price_total()
isFull=fn_Is_Cart_Full()

sCartText = ("<form name='Show_big_cart' action='"&Switch_Name&"cart_action.asp' method=post>"&_
    "<table border='0' width='100%'><tr>"&_
    "<td width='100%' class='big'><b>"&Shopping_Cart&"</b></td></tr></table>")

if isFull Then
    sCartText=sCartText & fn_create_cart(1,0,0,-1,-1,0,0)
    
    sCartText = sCartText & ("<br><br><TABLE width='427'><TR><TD width='147'>"&_
        fn_create_action_button ("Button_image_UpdateCart", "Re_Calculate", "Update Cart")&_
        "<input type=hidden name='Return_To' value='"&Return_To&"'></TD><TD width='210'>" &_
        fn_create_action_button ("Button_image_ContinueShopping", "Continue_Shopping", "Continue Shopping")&_
        "</TD><TD width='114'>" & fn_create_action_button ("Button_image_Checkout", "Check_Out", "Check Out")&_
        "</TD><TD width='68'>"&_
        "</TD><TD width='68' valign='top'></TD></TR></TABLE>"&_
        "<TABLE width='427'><TR><TD width='557'><table width='419'>")

    if When_Adding then
	    sCartText = sCartText & ("<tr><td colspan=3><HR></td></tr>"&_
	        "<tr><td align='middle' vAlign='top' width='27'>"&_
	        "<input name='Redirect_To_Cart_Old' type='hidden' value='"&sRedirect&"'>"&_
	        "<input CHECKED1 name='Redirect_To_Cart' type='radio' value='1'></td>"&_
	        "<td vAlign='top' colspan='2' width='378' class='normal'>When Adding Items to Cart, Show Shopping Cart</td></tr>"&_
	        "<tr><td align='middle' vAlign='top' width='27'><input CHECKED0 name='Redirect_To_Cart' type='radio' value='0'></td>"&_
	        "<td vAlign='top' colspan='2' width='378' class='normal'>When Adding Items to Cart, Hide Shopping Cart</td></tr>")
    else
	     sCartText = sCartText & ("<input name='Redirect_To_Cart' type='hidden' value='1'>"&_
	        "<input name='Redirect_To_Cart_Old' type='hidden' value='"&sRedirect&"'>"&_
	        "<TR><td vAlign='top' colspan='3'>Select update cart to keep cart changes</td>"&_
	        "</tr>")
    end if
    if Save_Cart then
	     sCartText = sCartText & ("<tr><td colspan=3><HR></td></tr>"&_
	        "<tr><td align='middle' vAlign='top' width='27'></td>"&_
	        "<td vAlign='top' width='205'>"&_
	        fn_create_action_button ("Button_image_SaveCart", "Save_Cart", "Save Cart")&_
	        "</td><td vAlign='top' width='167'>"&_
	        fn_create_action_button ("Button_image_RetrieveCart", "Retrieve_Cart", "Retrieve Cart")&_
	        "</td></tr>")
    end if
    sCartText = sCartText & ("</table></TD></TR></TABLE>")

else
    sCartText = sCartText & ("<BR><BR><B>Your shopping cart is empty</b><BR><BR>"&_
    "<input type=hidden name='Return_To' value='"&Return_To&"'>" &_
    fn_create_action_button("Button_image_ContinueShopping", "Continue_Shopping", "Continue Shopping"))
    if Save_Cart then
        sCartText = sCartText & (" " & fn_create_action_button ("Button_image_RetrieveCart", "Retrieve_Cart", "Retrieve Cart"))
    end if
end if

sCartText = sCartText & ("</form>")

sCartText = Replace (sCartText,"URL_STRING",url_string)
sCartText = Replace (sCartText,"CHECKED1",checked1)
sCartText = Replace (sCartText,"CHECKED0",checked0)

'PAYPAL EXPRESS CHECKOUT DISPLAY
if (Real_Time_Processor = 36 or PayPal_Express) and isFull  then 
    sCartText = sCartText & ("<form action='include/paypalPro/ExpressCheckout.asp' method=post>"&_
        "<input type='hidden' name= 'Auth_capture' value='"&auth_capture&"'>"&_
        "<input type='hidden' name='store_id' value='"&store_id&"'>"&_
        "<input type='hidden' name='shopper_id' value='"&Shopper_id&"'>"&_
        "<input type='hidden' name='total_amount' value='"&sTotal_cart&"'>"&_	
        "<TABLE width='427'><TR><td>"&_
        "<input type='image' id='submit' src='https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif'"&_
        " align='left' style='border:0;margin-right:7px;' alt='Checkout with PayPal'>"&_
        "<span style='font-size:11px; font-family: Arial, Verdana;'>"&_
        "Save time.  Checkout securely.  Pay without sharing your financial information.</span></td></tr></table>"&_	
	    "</Form>")
end if

response.Write sCartText


'Google CHECKOUT DISPLAY

%>
 <!-- #include file="include/google/googleglobal.asp" -->

<%


if GoogleCheckout and isFull then

sql_select = "exec wsp_settings_select "&store_id&";"
sub_write_log sql_select
rs_store.open sql_select,conn_store,1,1
if not rs_store.eof then
	GoogleCheckout_ButtonStyle=rs_Store("GoogleCheckout_ButtonStyle")
      end if
rs_store.close
if GoogleCheckout_ButtonStyle=0 then
Button_Style="white"
else
Button_Style="trans"
end if
DisplayButton "MEDIUM", True,Button_Style
end if
%>

<!--#include file="include/footer.asp"-->
