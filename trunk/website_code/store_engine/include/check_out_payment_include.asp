
<%
function max(x,y)
	if (x>y) then
		max=x
	else
		max=y
	end if
end function

function removeDupsArray(byVal sList)
  dim sNewList, aList, maxItems
  if not isNull(sList) then
    aList = split(sList,",")
  
    maxItems = ubound(aList)
    for x = 0 to (maxItems)
      if instr(sNewList, aList(x) & ",") <= 0 then sNewList = sNewList & aList(x) & ","
    next
    if sNewList<>"" then
       removeDupsArray = left(sNewList,len(sNewList)-1)
    end if

  else
    removeDupsArray = sList
  end if
end function

function fn_purchase_complete(oid)
    fn_purchase_unlock oid
    
end function

function fn_purchase_unlock(oid)
    sql_update = "exec wsp_purchase_unlock "&store_id&","&oid&";"
	fn_print_debug sql_update
	conn_store.Execute sql_update
end function

function fn_purchase_lock(oid,lock_timestamp,purchase_completed,verified)
	if purchase_completed then
		'order is already done dont need to process again
        fn_purchase_complete oid
	elseif datediff("n",lock_timestamp,now())<2 then
		'purchase was started, less then 2 minutes ago and hasnt been released, ie went to gateway but not completed
		fn_wait 10
		'wait 10 seconds then check if they are still there, if not abandon
		if response.isclientconnected=false then
			response.end
		end if
        'if visitor is still waiting check again to see if lock is released
          
		fn_redirect Switch_Name&"include/check_out_payment_action.asp"
    else
        sql_update = "exec wsp_purchase_lock "&store_id&","&oid&";"
    	fn_print_debug sql_update
    	conn_store.Execute sql_update
    end if
end function

'SEND AN EMAIL TO THE AFFILIATE, NOTIFING THAT A NEW ORDER
'WAS PLACED
sub affiliate_notify()

	sql_select = "select SA.Email from Store_Affiliates SA WITH (NOLOCK), Store_Purchases SP WITH (NOLOCK) where SA.store_id="&store_id&" and SA.Code=SP.CAME_FROM AND SP.oid="&oid&" and SA.Email_Notification<>0"
	fn_print_debug sql_select
	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

	if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		Email_Addr = mydata(myfields("email"),rowcounter)
		Email_Subject = "New order was placed at " &Store_Name
		Email_Body = "New order was placed at "&Store_Name&vbcrlf
		Email_Body = Email_Body&"Order is is "&oid&vbcrlf
		Email_Body = Email_Body&"Grand Total "&formatcurrency(grand_total)&vbcrlf&vbcrlf
		Email_Body = Email_Body&Store_Name&"You are being alerted of this because the order was placed via a link on your site."
		Send_Mail Store_email,Email_Addr,Email_Subject,Email_Body
	Next
	End if

	set myfields = Nothing
end sub

function fn_assign_pins(item_sku,quantity)
    sql_pin=""
    if Verified then
        sql_new = "select top "&quantity&" pin,other_info from Store_pin WITH (NOLOCK) where store_id="&store_id & " AND item_sku='" & item_sku & "' AND (pin_used is null or pin_used =0)"
        fn_print_debug sql_new
        set myfields=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_new,mydata,myfields,noRecords)

        if noRecords = 0 then
	        FOR rowcounter= 0 TO myfields("rowcount")
                sPin = mydata(myfields("pin"),rowcounter)
                sOther= mydata(myfields("other_info"),rowcounter)
                pininfo = sPin & "-" & sOther
                sql_pin = sql_pin & "exec wsp_pin_redeem "&store_id&","&cid&","&oid&",'"&item_sku&"','"&sPin&"';"
            Next
            sPins=myfields("rowcount")+1
        else 
            sPins=0
        end if
        if sPins<quantity then
            call pin_empty(item_sku,quantity-sPins)
        end if
            
        if pininfo<>"" then
            send_from = Session("Store_Email")
            if send_from="" or isnull(send_from) then
                send_from = Store_Email
            end if
        	
            send_cont = pin_buyer_email_body
            send_cont = send_cont&vbcrlf&"The pin code(s) you have purchased are below: "&pininfo&vbcrlf
            call Send_Mail_Html(send_from, shipemail, pin_email_Subject, send_cont)
        end if
    end if
    fn_assign_pins = sql_pin
End function

sub pin_empty(item_sku,sPinQuantity)
    pininfo="Store is out of pins for sku: "&item_sku&" please contact store owner."
    body = "Dear store owner, the StoreSecured system has sent you this email to alert you that the item with sku:"&item_sku&" has run out of pins.  We were unable to auto deliver "&sPinQuantity&" pin(s) for order id "&oid&" to email "&shipemail&"."
    call Send_Mail_Html(sNoReply_email, store_email, "Pins Empty in your store", body)
end sub

function fn_update_stock(item_id,item_sku,item_name,quantity,quantity_in_stock,quantity_control_number)
    sql_stock=""
    if (Verified AND Inventory_Reduce <>0) OR (not(verified) AND Unverified_Reduce <>0) then
        'update stock
        sql_stock=sql_stock&"exec wsp_item_qty_update "&store_id&","&item_id&","&oid&","&quantity&";"
        if Quantity_in_Stock>0 and (Quantity_in_Stock-Quantity <= Quantity_Control_Number or Quantity_in_Stock-Quantity <= 0) then
            Stock_Updated=-1
            subject="Out of stock items in your store"
            body = "Dear store owner, the StoreSecured system has sent you this email to alert you that the item with sku:"&item_sku&", and name:"&item_name&", has reached it's minimum sell quantity or quantity in stock has reached 0."
            call Send_Mail(Store_Email,Store_Email,subject,body)
        end if
    end if
    fn_update_stock=sql_stock
end function

function fn_assign_gift(gift_id,gift_amount,u_d_1,u_d_2,quantity)
    fn_print_debug "in assign gift"
    sql_gift=""
    for iQuantity=1 to formatnumber(quantity,0)
        Randomize
        GIFT_CODE = 1
		GIFT_CODE = "GCRT_"&Store_id&Oid&GIFT_CODE&cstr(Int((10000) * Rnd + lowerbound))
		GIFT_CODE = ucase(GIFT_CODE)
		GIFT_CODE = clearSpaces(GIFT_CODE)
		sql_gift = sql_gift & "exec wsp_giftcert_insert "&Store_Id&","&shopper_id&","&cid&",'"&shipemail&"',"&gift_id&","&oid&",'"&gift_code&"',"&gift_amount&",'"&u_d_1&"','"&u_d_2&"',"&verified&";"
        if verified then
		    from_mail = Session("Store_Email")
            if from_mail="" or isnull(from_mail) then
                from_mail = Store_Email
            end if

            send_cont = Gift_buyer_email_body
            if instr(send_cont,"%GIFT_CODE%")>0 then
                send_cont=replace(send_cont,"%GIFT_CODE%",gift_code)
            else
                send_cont = send_cont&vbcrlf&"Your gift certificate code is: "&gift_code
            end if
            send_cont=replace(send_cont,"%AMOUNT%",Currency_Format_Function(gift_amount))

            send_cont = send_cont&vbcrlf&u_d_2

            call Send_Mail_Html(from_mail, u_d_1, Gift_email_Subject, send_cont)
            send_cont="Here is a copy of the message sent to your gift certificate recipient.<BR>"&send_cont
            call Send_Mail_Html(from_mail, shipemail, Gift_email_Subject, send_cont)
		end if
    next
    fn_assign_gift = sql_gift
end function

' CLEAR SPACES IN THE RECEIVED STRING
function clearSpaces(theStr)
	rez = ""
	for i=1 to len(theStr)
		buf = mid(theStr,i,1)
		if (buf<>" ") then
			rez = rez&buf
		end if
	next
	clearSpaces = rez
end function




%>
