<!--#include file="include/header.asp"-->
<!--#include file="include/before_payment_include.asp"-->
<!--#include file="include/cart_display.asp"-->
<% 
Server.ScriptTimeout = 150
Dict_Special_Notes=Special_Notes
Dict_Shipping=Shipping
Dict_Coupon=Coupon

sShipMulti=0

if request.form("Coupon_Id") <>"" then
    Coupon_id = checkStringForQ(Request.Form("Coupon_Id"))
    fn_print_debug "gcrt="&instr(Coupon_Id,"GCRT_"&Store_Id)
	if instr(Coupon_Id,"GCRT_"&Store_Id)=1 then
		Giftcert_id=Coupon_id
		Coupon_id=""
	    Giftcert_Amounts = fn_Calc_Gift_Certificate(Giftcert_id)
	else
	    Coupon_Amounts = fn_Calc_Coupon_Discount(Coupon_id)
	end if
elseif fn_get_querystring("Remove_Coupon")<>"" then
    sRemove=fn_get_querystring("Remove_Coupon")
    if sRemove="Coupon" then
        sql_update = "exec wsp_coupon_remove "&Store_Id&","&shopper_id&";"
    elseif sRemove="Giftcert" then
        sql_update = "exec wsp_giftcert_remove "&Store_Id&","&shopper_id&";"
    end if
elseif fn_get_querystring("shipmulti")<>"" then
    sShipMulti=1

	    sql_update = sql_update&"exec wsp_trans_split "&Store_Id&","&Shopper_Id&";"
	    For Each Key In request.QueryString
            if instr(Key,"SAB_ShipAddr")=1 then
                SADDR=fn_get_querystring(Key)
                if (SADDR="") then
	                SADDR = 1
                elseif SADDR="New" then
	                fn_redirect Switch_Name&"Modify_my_Shipping.asp?ssadr=New"
                end if
                sTransaction_id=replace(Key,"SAB_ShipAddr_","")
                sql_update = sql_update & "exec wsp_trans_shipto_update "&Store_Id&","&Shopper_ID&","&SADDR&","&sTransaction_id&";"
            end if
        next
elseif fn_get_querystring("ChangeAddr")="Yes" then
        For Each Key In request.QueryString
            if instr(Key,"SAB_ShipAddr")=1 then
                SADDR=fn_get_querystring(Key)
                if (SADDR="") then
	                SADDR = 1
                elseif SADDR="New" then
	                fn_redirect Switch_Name&"Modify_my_Shipping.asp?ssadr=New"
                end if
            end if
        next
        if SADDR="" then
        	SADDR = 1
	   end if
	   sql_update = sql_update & "exec wsp_trans_shipto_update "&Store_Id&","&Shopper_ID&","&SADDR&",-1;"

else
    call sub_check_minimum
    sql_update = fn_recalc_cart()
    sql_update = sql_update&"exec wsp_purchase_create "&Store_Id&","&Shopper_Id&","&oid&","&startoid&";"
    fn_print_debug sql_update
    session("sql")=sql_update
    conn_store.execute sql_update
    sql_update=fn_Calc_Promotion_Discount()
end if

if sql_update<>"" then
    fn_print_debug sql_update
    conn_store.execute sql_update
end if

sPaymentText = ("<table border='0' width='100%' height='66'>")
sError=fn_get_querystring("Error")
if replace(replace(sError,",","")," ","") <> "" then
    sShippingErrorMessage=sError
    sPaymentText = sPaymentText & ("<TR><TD colspan=2 bgcolor=red><B><font color=FFFFFF>"&_
        Dict_Shipping&" Rates Error: "&sShippingErrorMessage&"</font></b><BR></td></tr>")
end if 

if Show_Coupon then
    sQuerystring=request.QueryString
    if sQuerystring<>"" then
        sQuerystring = replace(sQuerystring,"Remove_Coupon=Coupon&","")
        sQuerystring = replace(sQuerystring,"Remove_Coupon=GiftCert&","")
    end if
    fn_print_debug "querystring="&sQuerystring
    sPaymentText = sPaymentText & ("<form method=post action='"&Switch_Name&"before_payment.asp?"&sQuerystring&"'>"&_
	    "<tr><td height='19' class='normal' nowrap><b>"&Dict_Coupon&"</b></td>"&_
	    "<td height='19' align=left><input type='text' name='Coupon_Id' size='20'>"&_
	    "&nbsp;<input type='submit' name='Apply' value='Apply'></td></tr></form>")
	if coupon_id<>"" or giftcert_id<>"" then
	    if coupon_id<>"" then
	        sPaymentText = sPaymentText & ("<tr><td height='19' class='small'></td>"&_
	            "<td height='19' align=left class=small>Current Applied Coupon<BR>"&_
	            "Name: "&coupon_id&_
	            "  <a href='"&Switch_Name&"before_payment.asp?Remove_Coupon=Coupon&"&sQuerystring&"' "&_
	            "class=small>Remove</a></td></tr>")
	    end if
	    if giftcert_id<>"" then
    	    sPaymentText = sPaymentText & ("<tr><td height='19' class='small'></td>"&_
	            "<td height='19' align=left class=small>Current Applied Gift Certificate<BR>"&_
	            "Code: "&giftcert_id&", Amount: "&Currency_Format_Function(giftcert_amounts)&_
	            "  <a href='"&Switch_Name&"before_payment.asp?Remove_Coupon=GiftCert"&sQuerystring&"' "&_
	            "class=small>Remove</a></td></tr>")
	        sGCText = ("<tr><td></td><td class=small>The payment method chosen will be used in the event your gift certificate does not cover the order total.</td></tr>")
	    end if
	end if
	sPaymentText = sPaymentText & ("<tr><td colspan=2><HR></td></tr>")
else
    sGCText = ""
end if

sPaymentText = sPaymentText & ("<form method='POST' action='"&Switch_Name&"include/before_payment_action.asp' name='Before_Payment'>"&_
    "<tr><td width='80%' height='67' colspan='2' class='normal'>"&_
    "<b><em>Bill To</em></b>"&fn_display_cust_info(cid,0)&"</td></tr>"&_
    "<tr><td width='80%' height='19' colspan='2'><a href='Modify_my_Billing.asp' class='link'>"&_
    "Change Billing Address</a><br><br></td></tr></table>"&_
	"<table border='0' width='100%'>")

sPaymentText = sPaymentText & ("<tr><td colspan='2'>")
    
if not(Show_shipping) then 
    sPaymentText = sPaymentText & ("<input type='hidden' name='ShipAddr' value='0'>")
    sPaymentText = sPaymentText & (fn_other_options(sAppendString))
else
    sPaymentText = sPaymentText & ("<table border='0' width='100%' cellpadding='2' cellspacing='0'>"&_
        "<tr><td colspan='5'><a href='"&Switch_Name&"Modify_my_Shipping.asp' class='link'>"&_
	    "Change "&Dict_Shipping&" Address</a></td></tr>")
	  		
	sql_cart = "exec wsp_trans_shipping "&Store_Id&","&Shopper_Id&";"
    fn_print_debug sql_cart
    set myfieldsloc=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_cart,mydataloc,myfieldsloc,noRecordsloc)
    last_ship_location_id=""
    last_shipto=""
    fMulti_location=0
        
    if noRecordsloc = 0 then
        if myfieldsloc("rowcount")>0 then
            FOR rowcounterloc= 0 TO myfieldsloc("rowcount")
                ship_location_id=mydataloc(myfieldsloc("ship_location_id"),rowcounterloc)
                if (last_ship_location_id<>"" and ship_location_id<>last_ship_location_id) then
                    fMulti_location=1 
                    exit for       
                end if
                last_ship_location_id=ship_location_id
            NEXT
        end if
        last_ship_location_id=""
        last_shipto=""
        
        If sShipMulti=1 and (myfieldsloc("rowcount")>0 or (myfieldsloc("rowcount")=0 and mydataloc(myfieldsloc("quantity"),0)>1)) then
	        'USER WANTS A MULTI SHIPPING (EACH ITEM TO A DIFFERENT SHIPPING ADDRESS)	
	        sPaymentText = sPaymentText & ("<tr><td colspan='5' class='normal'><hr>"&_
		        "<input type='hidden' name='ShipOrderMulti' value='true'></td></tr>"&_
		        "<tr><td colspan='5' class='normal'><b><i>Each item shipped separately<BR>"&_
		        "<a class='link' href=""JavaScript:self.location='"&Switch_Name&"Before_Payment.asp';"">"&_
		        "Click here</a> to ship all items to a single location<br>&nbsp;</i></b></td></tr>")   
        else
            'USER WANTS A NORMAL SHIPPING (ALL ITEMS TO THE SAME ADDRESS)
	        sShipMulti=0
	        sPaymentText = sPaymentText & ("<tr><td colspan='5'><hr></td></tr>")
	        If (myfieldsloc("rowcount")>0 or (myfieldsloc("rowcount")=0 and mydataloc(myfieldsloc("quantity"),0)>1)) and (Ship_Multi or Ship_Multi="") then
	            sPaymentText = sPaymentText & ("<tr><td colspan='5' class='normal'>"&_
	                "<b><i>All Items Shipped to Single Address<BR>"&_
	                "<a class='link' href=""JavaScript:self.location='"&Switch_Name&"Before_Payment.asp?ShipMulti=TRUE';"" class='link'>"&_
	                "Click here</a> to ship items separately<br>&nbsp;</i></b></td></tr>")
	        End If
	    end if
	    
	    if fMulti_location=1 then
            'ITEMS ARE SHIPPING FROM DIFF LOCATIONS	
	        sPaymentText = sPaymentText & ("<tr><td colspan='5' class='normal'><HR>"&_
		        "Please note that your items will be shipped from multiple locations<HR></td></tr>")
        end if
	    
        FOR rowcounterloc= 0 TO myfieldsloc("rowcount")
            sTransaction_id=mydataloc(myfieldsloc("transaction_id"),rowcounterloc)
            ship_location_id=mydataloc(myfieldsloc("ship_location_id"),rowcounterloc)
            shipto=mydataloc(myfieldsloc("shipto"),rowcounterloc)
            sAppendStringShip="_"&shipto
            sAppendString="_"&sTransaction_id
	          
            if sShipMulti=1 or fMulti_location=1 then
                sPaymentText = sPaymentText & ("<tr><td>&nbsp;</td><td>&nbsp;</td>"&_
                    "<td colspan='3' bgcolor='#DDDDDD' class='normal'><b>"&_
                    mydataloc(myfieldsloc("quantity"),rowcounterloc)&" x "&_
                    mydataloc(myfieldsloc("item_name"),rowcounterloc)&"</b></font></td></tr>")
                if sShipMulti=1 then
                    sPaymentText = sPaymentText & ("<tr><td class='normal' colspan=4><B>Select "&Dict_Shipping&" Address</b></td>"&_
                        "<td width='50%'>"&fn_displayAddressBook("ShipAddr"&sAppendString,shipto)&"</td></tr>")
            
                    sAddrBookParms=sAddrBookParms&_
		                "for (i=0;i<document.forms['Before_Payment']."&"ShipAddr"&sAppendString&".length;i++)"&vbcrlf&_
	                    "if (document.forms['Before_Payment']."&"ShipAddr"&sAppendString&"[i].selected)"&vbcrlf&_
	                    "parmsString = parmsString + '&ChangeAddr=Yes&SAB_"&"ShipAddr"&sAppendString&"=' + document.forms['Before_Payment']."&"ShipAddr"&sAppendString&"[i].value;"&vbcrlf
                end if
            end if
            
            If ship_location_id<>last_ship_location_id or shipto<>last_shipto then
                sAppendString="_"&ship_location_id&"_"&shipto
        			if sShipMulti<>1 then
				   sPaymentText = sPaymentText & ("<tr><td class='normal' colspan=4><B>Select "&Dict_Shipping&" Address</b></td>"&_
                        "<td width='50%'>"&fn_displayAddressBook("ShipAddr"&sAppendString,shipto)&"</td></tr>")
            
                    sAddrBookParms=sAddrBookParms&_
		                "for (i=0;i<document.forms['Before_Payment']."&"ShipAddr"&sAppendString&".length;i++)"&vbcrlf&_
	                    "if (document.forms['Before_Payment']."&"ShipAddr"&sAppendString&"[i].selected)"&vbcrlf&_
	                    "parmsString = parmsString + '&ChangeAddr=Yes&SAB_"&"ShipAddr"&sAppendString&"=' + document.forms['Before_Payment']."&"ShipAddr"&sAppendString&"[i].value;"&vbcrlf
				end if
				sPaymentText = sPaymentText & ("<tr><td class='normal' colspan=4>"&_
                    "<B>Select "&Dict_Shipping&" Method</b></td><td width='50%' class='normal'>")
                sel_name="Shipping_Method_name_Price"&sAppendString
                sPaymentText = sPaymentText & (fn_Shipping_Options_Multi (sel_name, shipto, ship_location_id))
                sExtraValidation=sExtraValidation&"frmvalidator.addValidation("""&sel_name&""",""req"",""Please select a "&Dict_Shipping&" method."");"
                sPaymentText = sPaymentText & ("</td></tr><tr>")
		        sItem_Id=mydataloc(myfieldsloc("item_id"),rowcounterloc)
		        
		        sPaymentText = sPaymentText & (fn_other_options(sAppendString))
	        End If
	        last_ship_location_id=ship_location_id
	        last_shipto=shipto
        Next
    end if
    
    	
    
    sAddrBookJsText = ("<script language=""javaScript"">"&vbcrlf&_
        "function changeAddressBookSelection()"&vbcrlf&_
	    "{var parmsString;parmsString="""";"&vbcrlf&_
	    sAddrBookParms&_
	    "self.location='"&Switch_Name&"Before_Payment.asp?ShipMulti="&fn_get_querystring("ShipMulti")&"'+parmsString;}"&vbcrlf&_
        "</script>"&vbcrlf)

    sPaymentText = sPaymentText & ("<tr><td colspan='5'><hr></td></tr></table>"&sAddrBookJsText)
    set myfieldsloc = Nothing
End If
sPaymentText = sPaymentText & ("</td></tr>")
 
if Enable_Rewards then
    if Reward_left => Rewards_Minimum and Reward_left > 0 and Enable_Rewards then
        sPaymentText = sPaymentText & ("<tr><td width='50%' class='normal'><b>Use Rewards Balance</b></td>"&_
	        "<td width='50%' class='normal'><input type='radio' value='-1' name='Rewards'>Yes ("&_
	        Currency_Format_Function(Reward_left)&" Credit)<br>"&_
            "<input type='radio' value='0' name='Rewards' checked>No</td></tr>")
    end if
end if
	
if ((Real_Time_Processor = 36 or Paypal_Express) and len(trim(session("token")))>0) then
    sPaymentText = sPaymentText & ("<input type=hidden name=payment_method value='PayPal-ExpressCheckout'>"&_
        "<input type=hidden name=paypal_token  value='"&fn_get_querystring("token")&"'>")
else
    sPaymentText = sPaymentText & ("<tr><td width='50%' class='normal'><b>Select Payment</b></td>"&_
	    "<td width='50%' class='normal'>"&fn_disp_payment()&"</td></tr>"&sGCText)
end if

sPaymentText = sPaymentText & ("<TR><TD colspan=2 align=center><BR>"&_
    fn_create_action_button ("Button_image_Continue", "Before_Payment_Continue", "Continue")&_
	"</TD></TR>")
sPaymentText = sPaymentText & ("</TABLE><input type=hidden name=Form_Name value=Before_Payment_Continue></form>")

sFormName = "Before_Payment"
sSubmitName = "Before_Payment_Continue"

sPaymentText = sPaymentText & (vbcrlf&"<SCRIPT language=""JavaScript"">"&vbcrlf&_
    "var frmvalidator  = new Validator("""&sFormName&""");"&vbcrlf&_
    "frmvalidator.addValidation(""Payment_Method"",""req"",""Please select a payment method."");"&vbcrlf&_
    sExtraValidation&vbcrlf&"</script>"&vbcrlf)
    
response.Write sPaymentText
%>
<!--#include file="include/footer.asp"-->
