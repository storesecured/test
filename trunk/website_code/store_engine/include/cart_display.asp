<%
function fn_create_cart (iEditable,iCompleted,isEmail,ShipTo,Ship_Location_Id,this_oid,verified)
    if iCompleted=0 then
        'POPULATE RECORDS FROM STORE_SHOPPINGCART TABLE
        sql_Show_big_cart = "exec wsp_cart_display "&Store_Id&","&Shopper_ID&";"
    else
        'POPULATE RECORDS FROM STORE_TRANSACTIONS TABLE
        if this_oid="" or this_oid=Null or not(isNumeric(this_oid)) then
        	this_oid=0
     	end if

        sql_Show_big_cart = "exec wsp_trans_display "&Store_Id&","&Shopper_ID&","&this_oid&";"

    end if
    fn_print_debug sql_Show_big_cart
    set myfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_Show_big_cart,mydata,myfields,noRecords)

    sCreateCart=""
    sEmptyCol = "<td class='normal' align=right>&nbsp;</td>"
    
    Line_id = 1
	if noRecords = 0 then
        Sale_Price_Total=0
        Recurring_Total=0
        iPromotions_exist=0
        iNotes_exist=0
      
        FOR rowcounter= 0 TO myfields("rowcount")
            Quantity = mydata(myfields("quantity"),rowcounter)
		    Sale_Price = mydata(myfields("sale_price"),rowcounter)
		    Cust_price = mydata(myfields("cust_price"),rowcounter)
		    Retail_Price = Currency_Format_Function(mydata(myfields("retail_price"),rowcounter))
		    Item_Id = mydata(myfields("item_id"),rowcounter)
		    Item_Page_Name = mydata(myfields("item_page_name"),rowcounter)
		    Full_Name = mydata(myfields("full_name"),rowcounter)
		    Custom_Link = mydata(myfields("custom_link"),rowcounter)
		    Item_name = mydata(myfields("item_name"),rowcounter)
		    Shopping_cart_id = mydata(myfields("shopping_cart_id"),rowcounter)
		    Item_Attribute_ids = mydata(myfields("item_attribute_ids"),rowcounter)
		    Recurring_Fee = mydata(myfields("recurring_fee"),rowcounter)
		    Recurring_Days_Initial = mydata(myfields("recurring_days"),rowcounter)
		    ImageS_Path = mydata(myfields("images_path"),rowcounter)
		    sCreateCart = sCreateCart & ("<tr>")
		    Ext_Sale_Price=Sale_Price*Quantity
		    Sale_Price_Total=Sale_Price_Total+Ext_Sale_Price
		    
            if Ext_Sale_Price = 0 then
			    Ext_Sale_Price = "<i>Free</i>"
		    else
			    Ext_Sale_Price = Currency_Format_Function(Sale_Price*Quantity)
		    End if
		    
		    if Sale_Price = 0 then
			    Sale_Price = "<i>Free</i>"
		    else
			    Sale_Price = Currency_Format_Function(mydata(myfields("sale_price"),rowcounter))
			    Sale_Price_unformatted = mydata(myfields("sale_price"),rowcounter)
		    End if
			if Cart_Thumbnails then
				sCreateCart = sCreateCart & ("<td>" & fn_create_image(images_Path,Item_name) & "</td>")
			End If
			if iEditable=1 then
			    if Hide_Quantity=1 then
			        sQuantityOutput="<INPUT type='hidden' name='Quantity_"&Line_id&"' value='"&Quantity&"'>"    
			    else
			        sQuantityOutput="<td align='center'><INPUT size='3' name='Quantity_"&Line_id&"' value='"&Quantity&"' "&_
			            "onKeyPress=""return goodchars(event,'0123456789.')""></td>"
			    end if
			    sQuantityOutput=sQuantityOutput&"<INPUT type=hidden name='Quantity_Old_"&Line_id&"' value='"&Quantity&"'>"
  			else
  			    sQuantityOutput="<td align='center'>"&Quantity&"</td>"
  			end if      
		    sCreateCart = sCreateCart & sQuantityOutput
      		u_d_1 = mydata(myfields("u_d_1"),rowcounter)
			u_d_2 = mydata(myfields("u_d_2"),rowcounter)
			u_d_3 = mydata(myfields("u_d_3"),rowcounter)
			u_d_4 = mydata(myfields("u_d_4"),rowcounter)
			u_d_5 = mydata(myfields("u_d_5"),rowcounter)
			u_d_1_name = mydata(myfields("u_d_1_name"),rowcounter)
			u_d_2_name = mydata(myfields("u_d_2_name"),rowcounter)
			u_d_3_name = mydata(myfields("u_d_3_name"),rowcounter)
			u_d_4_name = mydata(myfields("u_d_4_name"),rowcounter)
			u_d_5_name = mydata(myfields("u_d_5_name"),rowcounter)
			Promotion_ids = mydata(myfields("promotion_ids"),rowcounter)
			Notes = mydata(myfields("notes"),rowcounter)
			File_Location = mydata(myfields("file_location"),rowcounter)
            sItemQueryString=""
            if iCompleted=1 then
                sLink=""
            elseif Custom_Link <> "" then
                sLink=Custom_Link
            elseif Item_ID<0 then
                sLink=""
            else
                sLink = fn_item_url(full_name,item_page_name)
                if u_d_1 <> "" and not isNull(u_d_1) then
                    sItemQueryString = sItemQueryString&("&u_d_1_"&Item_Id&"="&server.urlencode(u_d_1))
                end if
                if u_d_2 <> "" and not isNull(u_d_2) then
                    sItemQueryString = sItemQueryString&("&u_d_2_"&Item_Id&"="&server.urlencode(u_d_2))
                end if
                if u_d_3 <> "" and not isNull(u_d_3) then
                    sItemQueryString = sItemQueryString&("&u_d_3_"&Item_Id&"="&server.urlencode(u_d_3))
                end if
                if u_d_4 <> "" and not isNull(u_d_4) then
                    sItemQueryString = sItemQueryString&("&u_d_4_"&Item_Id&"="&server.urlencode(u_d_4))
                end if
                if u_d_5 <> "" and not isNull(u_d_5) then
                    sItemQueryString = sItemQueryString&("&u_d_5_"&Item_Id&"="&server.urlencode(u_d_5))
                end if
                if Item_Attribute_ids <> "" and not isNull(Item_Attribute_ids) then
                    sItemQueryString = sItemQueryString&("&Attributes="&server.urlencode(Item_Attribute_ids))
                end if
                if sItemQueryString<>"" then
                    sLen=len(sItemQueryString)
                    sLink = sLink & ("?"&right(sItemQueryString,sLen-1))
                end if    
            end if
            if sLink<>"" then
                sCreateCart = sCreateCart & ("<td><a href='"&sLink&"' class='link'>"&Item_name&"</a>")
            else
                sCreateCart = sCreateCart & ("<td>"&Item_name)
            end if
            sCreateCartUd=""
            sCreateCartUd = sCreateCartUd & fn_create_ud(User_Defined_Fields,u_d_1_name,u_d_1,"u_d_1",Line_id,iEditable)
			sCreateCartUd = sCreateCartUd & fn_create_ud(User_Defined_Fields_2,u_d_2_name,u_d_2,"u_d_2",Line_id,iEditable)
			sCreateCartUd = sCreateCartUd & fn_create_ud(User_Defined_Fields_3,u_d_3_name,u_d_3,"u_d_3",Line_id,iEditable)
			sCreateCartUd = sCreateCartUd & fn_create_ud(User_Defined_Fields_4,u_d_4_name,u_d_4,"u_d_4",Line_id,iEditable)
			sCreateCartUd = sCreateCartUd & fn_create_ud(User_Defined_Fields_5,u_d_5_name,u_d_5,"u_d_5",Line_id,iEditable)
			if sCreateCartUd<>"" then
			    sCreateCart = sCreateCart & ("<br><table>"&sCreateCartUd&"</table>")
			end if 
			fn_print_debug "user_id="&user_id
               fn_print_debug "isEmail="&isEmail
               fn_print_debug "EXPRESS_CHECKOUT_CUSTOMER="&EXPRESS_CHECKOUT_CUSTOMER
			if isEmail=0 then
				fn_print_debug "email is equal=0"
			end if
			if user_id<>EXPRESS_CHECKOUT_CUSTOMER then
			 	fn_print_debug "user id<>"
			end if
			if iCompleted=1 and Verified and File_Location <> "" and File_Location <> "0" and (isEmail=0 or (user_id<>EXPRESS_CHECKOUT_CUSTOMER)) then
		            if Instr(File_Location,"http://") > 0 then
				        sFilename = File_Location
			        else
				        sFilename = Switch_Name&"include/esd_download.asp?File_Location="&server.urlencode(File_Location)
			        end if
			        sCreateCart = sCreateCart & ("<BR><a href='"&_
		                sFilename&"' class='link'>Download</a>")
	            End If
			sCreateCart = sCreateCart & ("</td>")  
			if not(Hide_Retail_Price) then
		        sCreateCart = sCreateCart & ("<td class='normal' align=right>"&Retail_Price&"</td>")
			end if
            sCreateCart = sCreateCart & ("<td class='normal' align=right>"&Sale_Price&"</td>")
			sCreateCart = sCreateCart & ("<td class='normal' align=right>"&Ext_Sale_Price&"</td>")
			if Recurring_Fee <> 0 then
			    Recurring_Total=Recurring_Total+Recurring_Fee
				sCreateCart = sCreateCart & ("<td class='normal' align=right>"&Currency_Format_Function(Recurring_Fee)&"</td>")
			else
			    sCreateCart = sCreateCart & ("%OBJ_RECURRING_OBJ%")    
			end if
			if iEditable=1 then
			    sCreateCart = sCreateCart & ("<td>"&_
			        "<input type='checkbox' name='Remove_From_Cart_"&Line_id&"' value='ON'>"&_
			        "<input type='hidden' name='Item_Id_"&Line_id&"' value='"&Item_Id&"'>"&_
			        "<input type='hidden' name='Cust_Price_"&Line_id&"' value='"&Cust_Price&"'>"&_
			        "<input type='hidden' name='Shopping_cart_id_"&Line_id&"' value='"&Shopping_cart_id&"'>"&_
			        "<input type='hidden' name='Item_Attribute_ids_"&Line_id&"' value='"&Item_Attribute_ids&"'>"&_
			        "<input type='hidden' name='Sale_Price_"&Line_id&"' value='"&Sale_Price_unformatted&"'>"&_
			        "</td>")
			else
			    if promotion_ids<>"" then
			        iPromotions_exist=1
                    sCreateCart = sCreateCart & ("<td class='small' align=center><i>Promotion</i></td>")
                else
                    sCreateCart = sCreateCart & ("%OBJ_PROMOTION_OBJ%")    
                end if
                if notes<>"" then
                    iNotes_exist=1
                    sCreateCart = sCreateCart & ("<td class='small' align=center nowrap>"&Notes&"</td>")
                else
                    sCreateCart = sCreateCart & ("%OBJ_NOTES_OBJ%") 
                end if
                
			end if

            sCreateCart = sCreateCart & ("</tr>"&sLineBreak)
		    Line_id = Line_id + 1
	    Next
	    
	    sCreateCartHead=""
        sCreateCartHead = sCreateCartHead & ("<table cellspacing=0 cellpadding=0 border=0 width='100%'><tr><td>"&_
            "<table cellspacing='0' cellpadding='3' border='1' width='100%'><tr>")
        if Cart_Thumbnails then
            sCreateCartHead = sCreateCartHead & ("<td class='normal'><STRONG>Image</STRONG></td>")
        End If
        if Hide_Quantity =0 then
            sCreateCartHead = sCreateCartHead & ("<td class='normal'><STRONG>Quantity</STRONG></td>")
        end if
        sCreateCartHead = sCreateCartHead & ("<td class='normal' width='80%'><STRONG>Name</STRONG></td>")
        if not(Hide_Retail_Price) then
            sCreateCartHead = sCreateCartHead & ("<td class='normal'><STRONG>Retail Price</STRONG></FONT></td>")
        end if
        sCreateCartHead = sCreateCartHead & ("<td class='normal'><STRONG>Your Price</STRONG></FONT></td>"&_
            "<td class='normal'><STRONG>Total</STRONG></td>")
        if recurring_Total <> 0 then
            sCreateCartHead = sCreateCartHead & ("<td class='normal'><STRONG>Recurring</STRONG></td>")
            sCreateCart = fn_replace(sCreateCart,"%OBJ_RECURRING_OBJ%",sEmptyCol)
        else
            sCreateCart = fn_replace(sCreateCart,"%OBJ_RECURRING_OBJ%","")
        end if

        if iEditable=1 then
            sCreateCartHead = sCreateCartHead & ("<td class='normal'><strong>Remove</strong></td>")
        else
            if iPromotions_exist=1 then
                sCreateCartHead = sCreateCartHead & (sEmptyCol)
                sCreateCart = fn_replace(sCreateCart,"%OBJ_PROMOTION_OBJ%",sEmptyCol)
            else
                sCreateCart = fn_replace(sCreateCart,"%OBJ_PROMOTION_OBJ%","")
            end if
            if iNotes_exist=1 then
                sCreateCartHead = sCreateCartHead & (sEmptyCol)
                sCreateCart = fn_replace(sCreateCart,"%OBJ_NOTES_OBJ%",sEmptyCol)
            else
                sCreateCart = fn_replace(sCreateCart,"%OBJ_NOTES_OBJ%","")
            end if
        end if
        sCreateCartHead=sCreateCartHead&("</tr>"&sLineBreak)
	    sCreateCart = sCreateCartHead & sCreateCart

		sCreateCart = sCreateCart & ("</table>")
		
		if iEditable=1 then	
		    sCreateCart = sCreateCart&("<BR><table cellspacing='0' cellpadding='0' border='0' align='right' width='50%'>"&_
		        "<tr><td align='right' class='normal'><STRONG>Sub Total</STRONG></td>"&_
                "<td align='right' class='normal'><STRONG>"&Currency_Format_Function(Sale_Price_Total)&"</STRONG></td>")
            if recurring_Total > 0 then
                sCreateCart = sCreateCart & ("</tr><tr>"&_
                    "<td align='right' class='normal'><STRONG>Every "&Recurring_days_initial&" days</STRONG></td>"&_
                    "<td align='right' class='normal'><STRONG>"&Currency_Format_Function(Recurring_Total)&"</STRONG></td>")
            end if
            sCreateCart = sCreateCart & ("</tr></table>")
        end if
        sCreateCart = sCreateCart & ("</td></tr></table>")
    End if
    set myfields = Nothing

    fn_create_cart=sCreateCart
end function

function fn_display_invoice (coid,shopper_id,iComplete,isEmail)
	fn_print_debug "isEmail="&isEmail
    if iComplete=1 then
        sql_Invoice_Custom = "Select Invoice_header, Invoice_footer From Store_Settings WITH (NOLOCK) where Store_id = "&Store_id
        fn_print_debug sql_Invoice_Custom
        session("sql")=sql_Invoice_Custom
	   rs_Store.open sql_Invoice_Custom,conn_store,1,1
        if not rs_Store.eof then
	        Invoice_header = Rs_store("Invoice_header")
	        Invoice_footer = Rs_store("Invoice_footer")
        end if
        rs_Store.Close
    else
        Invoice_header = ""
	    Invoice_footer = ""
    end if
    if coid="" then
       coid=0
    end if
    if coid=0 then
       sql_select_purchases = "Select * from Store_Purchases WITH (NOLOCK) where Store_id ="&Store_id&" and Shopper_id ="&Shopper_ID&" order by sys_created desc"
    elseif isEmail=0 then
       sql_select_purchases = "Select * from Store_Purchases WITH (NOLOCK) where Store_id ="&Store_id&" and (oid = "&coid&" or masteroid= "&coid&") order by oid"
    else
       sql_select_purchases = "Select * from Store_Purchases WITH (NOLOCK) where Store_id ="&Store_id&" and oid = "&coid
    end if
    fn_print_debug sql_select_purchases
    set payfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select_purchases,paydata,payfields,noPayRecords)
    if noPayRecords = 1 then
        set payfields=Nothing
        fn_redirect Switch_Name&"Show_Big_Cart.asp"
    end if
    sOrderText=""
    Display_Grand_Total=0
    GGrand_Total=0
    FOR rowcounterPay= 0 TO payfields("rowcount")
        if coid=0 then
            coid=paydata(payfields("oid"),rowcounterPay)
            Session("oid")=coid
        end if
        this_oid=paydata(payfields("oid"),rowcounterPay)
        Shipping_Method_Price = paydata(payfields("shipping_method_price"),rowcounterPay)
        Shipping_Method_Name = paydata(payfields("shipping_method_name"),rowcounterPay)
        Tax = paydata(payfields("tax"),rowcounterPay)
        Recurring_Total = paydata(payfields("recurring_total"),rowcounterPay)
        Recurring_FeeT = paydata(payfields("recurring_fee"),rowcounterPay)
        Recurring_Tax = paydata(payfields("recurring_tax"),rowcounterPay)
        Recurring_Days = paydata(payfields("recurring_days"),rowcounterPay)
        Coupon_id = paydata(payfields("coupon_id"),rowcounterPay)
        Coupon_Amount = paydata(payfields("coupon_amount"),rowcounterPay)
        giftcert_id = paydata(payfields("giftcert_id"),rowcounterPay)
        giftcert_Amount = paydata(payfields("giftcert_amount"),rowcounterPay)
        rewards_Amount = paydata(payfields("rewards"),rowcounterPay)
        Gift_Wrapping = paydata(payfields("gift_wrapping"),rowcounterPay)
        Gift_Wrapping_Amount = paydata(payfields("gift_wrapping_amount"),rowcounterPay)
        Total = paydata(payfields("total"),rowcounterPay)
        Grand_Total = paydata(payfields("grand_total"),rowcounterPay)
        Payment_Method = paydata(payfields("payment_method"),rowcounterPay)
        cust_notes = paydata(payfields("cust_notes"),rowcounterPay)
        custom_fields = paydata(payfields("custom_fields"),rowcounterPay)
        gift_message = paydata(payfields("gift_message"),rowcounterPay)
        ShipTo = paydata(payfields("shipto"),rowcounterPay)
        Ship_Location_Id = paydata(payfields("ship_location_id"),rowcounterPay)
        purchase_date = paydata(payfields("purchase_date"),rowcounterPay)
        purchase_completed = paydata(payfields("purchase_completed"),rowcounterPay)
        verified=paydata(payfields("verified"),rowcounterPay)
        cid = paydata(payfields("cid"),rowcounterPay)
        ccid = paydata(payfields("ccid"),rowcounterPay)
        Cust_PO = paydata(payfields("cust_po"),rowcounterPay)
        if purchase_completed<>0 then
	   	Payment_Method_Message=fn_payment_message(Payment_Method)
        else
	   	Payment_Method_Message=""
	   end if
	   if Invoice_header<>"" then
            sOrderText = sOrderText&"<HR>"&Invoice_header
        end if
        fn_print_debug "this verified="&verified
        sOrderText = sOrderText&("<HR><font class='normal'>"&fn_display_address(1)&"</font>"&_
		    "<HR><table width='100%' border=0 cellpadding=2 cellspacing=0>"&_
		    "<tr><td><table width='100%' border=0 cellpadding=2 cellspacing=0>"&_
			"<tr><td class='normal'><B>Date </B>"&formatdatetime(purchase_date,2)&"</td>"&sLineBreak&_
			"<td class='normal'><B>Contact </b>"&first_name&"&nbsp;"&last_name&"</td>"&sLineBreak&_
			"<td class='normal'><B>Customer ID </B>"&CCid&"</td></tr>"&sLineBreak&_
	        "<tr><td class='normal'><B>Order ID </B>"&this_oid&"</td>"&sLineBreak&_
	        "<td class='normal'><B>Terms </B>"&Payment_Method&"</td>"&sLineBreak&_
	        "<td class='normal'>")
        if Cust_PO<>"" then
            sOrderText = sOrderText&("<b>Purchase Order "&Cust_PO&"</b>"&sLineBreak)
        end if
        sOrderText = sOrderText&("</td></tr></table>"&_
            "<hr width='100%' align=left>"&_
            "<table width='100%' border=0 cellpadding=2 cellspacing=0>"&_
	        "<tr><td valign=top width='50%' class='normal'>"&_
		    "<B>Billing</B>"&fn_display_cust_info(cid,0)&"</TD>"&sLineBreak&_
            "<td width='50%' valign=top class='normal'>")    
        if Show_shipping then
            sOrderText = sOrderText&("<B>"&Shipping&"</B>"&sLineBreak&_
                fn_display_cust_info(cid,ShipTo))
        end if 
	    sOrderText = sOrderText&("</td></tr></table>")

        sOrderText=sOrderText & fn_create_cart(0,iComplete,isEmail,ShipTo,Ship_Location_Id,this_oid,verified)
        sOrderText=sOrderText & ("<br>"&_
            "<table cellspacing='0' cellpadding='5' border='1' align='right' width='70%'>"&_
            "<tr><td align='left' class='normal'>"&_
            "Sub Total</td>"&_
            "<td align='right' class='normal'>"&_
            Currency_Format_Function(Total)&"</td></tr>"&sLineBreak)
        If Coupon_Id <> "" then
	        sOrderText=sOrderText & ("<tr><td align='left' class='normal'>"&_
	            "<i>Coupon ("&Coupon_Id&")</i></td>"&_
		        "<td align='right' class='normal'><i>-"&_
		        Currency_Format_Function(Coupon_Amount)&"</i></td></tr>"&sLineBreak)
        End If
        If cdbl(rewards_Amount) > 0 then
	        sOrderText=sOrderText & ("<tr><td align='left' class='normal'>"&_
	            "<i>Rewards</i></td>"&_
		        "<td align='right' class='normal'><i>-"&_
		        Currency_Format_Function(rewards_Amount)&"</i></td></tr>"&sLineBreak)
        End If
        If Gift_Wrapping <> 0 then
	        sOrderText=sOrderText & ("<tr><td align='left' class='normal'>"&_
	            "Gift Wrapping</td>"&_
		        "<td align='right' class='normal'>"&_
		        Currency_Format_Function(Gift_Wrapping_amount)&"</td></tr>"&sLineBreak)
        End If
        if Show_shipping then
            sOrderText=sOrderText & ("<tr><td align='left' class='normal'>"&_
                "Shipping ("&Shipping_Method_Name&")</td>"&_
	            "<td align='right' class='normal'>"&_
	            Currency_Format_Function(Shipping_Method_price)&"</td></tr>"&sLineBreak)
        end if
        if Tax <>0 then
            sOrderText=sOrderText & ("<tr><td align='left' class='normal'>Tax</td>"&_
	            "<td align='right' class='normal'>"&Currency_Format_Function(Tax)&""&_
	            "</td></tr>"&sLineBreak)
        end if
        If Giftcert_Id <> "" and cdbl(Giftcert_Amount)>0 then
	        sOrderText=sOrderText & ("<tr><td align='left' class='normal'>"&_
	            "<i>Gift Certificate</i></td>"&_
		        "<td align='right' class='normal'><i>-"&_
		        Currency_Format_Function(Giftcert_Amount)&"</i></td></tr>"&sLineBreak)
        End If
        if payfields("rowcount") = 0 then
            total_name="Grand Total"
        else
            total_name="Total"
        end if
        sOrderText=sOrderText & ("<tr><td align='left' class='normal'>"&_
            "<STRONG>"&total_name&"</STRONG></td><td align='right' class='normal' align='right'>"&_
            "<STRONG>"&Currency_Format_Function(Grand_Total)&"</STRONG></td></tr>"&sLineBreak)
	    Display_Grand_Total = Display_Grand_Total + Grand_Total
	    GGrand_Total=GGrand_Total+Grand_Total
        if Recurring_Total <> 0 then
            sOrderText=sOrderText & ("<tr>"&_
                "<td align='left' class='normal'>Recurring Total</td>"&_
                "<td align='right' class='normal' align=right>"&_
                Currency_Format_Function(Recurring_FeeT)&"</td></tr>"&sLineBreak)
            if Recurring_Tax>0 then
                sOrderText=sOrderText & ("<tr><td align='left' class='normal'>Recurring Tax</td>"&_
                	"<td align='right' class='normal' align=right>"&_
                	Currency_Format_Function(Recurring_Tax)&"</td></tr>"&sLineBreak)
            end if
            sOrderText=sOrderText & ("<tr><td align='left' class='normal'><STRONG>Every "&Recurring_Days&" days</STRONG></td>"&_
                "<td align='right' class='normal' align=right><STRONG>"&_
                Currency_Format_Function(Recurring_Total)&"</STRONG></td></tr>"&sLineBreak)
        end if
        sOrderText=sOrderText & ("</table></td></tr></table><br>")
        if iComplete=1 then
            iNotes=0
            sNotesText=("<HR><table width='75%' border=0 cellspacing=1 cellpadding=1 height='1'>")
            if Payment_Method_Message<>"" then
		        iNotes=1
		        sNotesText=sNotesText & ("<TR><TD height='1' class='normal'>"&_
                    "<!-- start payment method message -->"&Payment_Method_Message&"<!-- end payment method message -->"&_
                    "</TD></TR>"&sLineBreak)
            end if
            if Cust_Notes<>"" then
		        iNotes=1
		        sNotesText=sNotesText & ("<TR><TD height='1' class='normal'>"&_
                    "<!-- start cust notes -->"&Cust_Notes&"<!-- end cust notes -->"&_
                    "</TD></TR>"&sLineBreak)
            end if
            if custom_fields<>"" then
		        iNotes=1
		        sNotesText=sNotesText & ("<TR><TD height='1' class='normal'>"&_
                    "<!-- start custom fields -->"&custom_fields&"<!-- end custom fields -->"&_
                    "</TD></TR>"&sLineBreak)
            end if
            if gift_message<>"" then
		        iNotes=1
		        sNotesText=sNotesText & ("<TR><TD height='1' class='normal'>"&_
                    "<!-- start gift message -->"&gift_message&"<!-- end gift message -->"&_
                    "</TD></TR>"&sLineBreak)
            end if
            sNotesText=sNotesText & ("</table>")
            if iNotes=1 then
                sOrderText=sOrderText&sNotesText
            end if
        end if
        if Invoice_footer<>"" then
            sOrderText = sOrderText&Invoice_footer
        end if
    next
    
    if payfields("rowcount")>0 then
        sOrderText = sOrderText & ("<HR><font class='normal'><STRONG>Grand Total: "&Currency_Format_Function(Display_Grand_Total)&"</strong></font><hr>"&sLineBreak)
    end if
    set payfields=Nothing
    fn_display_invoice=sOrderText
end function

function fn_display_cust_info(cid,Record_type)
	' RECORD_TYPE PARAM : 0=BILLING ADDRESS; 1=SHIPPING ADDRESS 1; 2=SHIPPING_ADDRESS 2; ETC
	sql_select_cust =  "exec wsp_customer_lookup "&store_id&","&cid&","&record_type&";"
	fn_print_debug sql_select_cust
	session("sql") = sql_select_cust
	set rs_store = conn_store.execute (sql_select_cust)
	on error resume next
    if rs_Store("State") = "AA" then
        State = ""
    else
        State = rs_Store("State") & " "
    end if
    if Record_Type<>0 and not(rs_Store("Is_Residential")) then
        Residential="<BR>Residential: No"
    else
        Residential=""
    end if

    Company = rs_Store("Company")
    Address2 = rs_Store("Address2")
    Fax = rs_Store("Fax")
    sCustText = "<blockquote><font class='normal'>"
    if Company<>"" then
        sCustText = sCustText & ("<b>"&Company&"</b><br>")
    end if
    sCustText = sCustText & (rs_Store("First_name")&"&nbsp;"&rs_Store("Last_name")&"<br>"&_
        rs_Store("Address1")&"<br>")
    if Address2 <> "" then
        sCustText = sCustText & (Address2&"<br>")
    end if
    sCustText = sCustText & (rs_Store("City")&", "&State&" "&rs_Store("zip")&"<br>"&_
        rs_Store("Country")&"<br>Phone: "&rs_Store("Phone")&"<br>")
    if Fax <> "" then
        sCustText = sCustText & ("Fax: "&Fax&"<br>")
    end if
    sCustText = sCustText & (rs_Store("EMail")&Residential&"</font></blockquote>")
  
    rs_Store.Close
    fn_display_cust_info = sCustText
End function

function fn_create_ud(User_Defined_Fields_Format,u_d_name,u_d_value,u_d_number,Line_id,iEditable)
    sUDCart=""
    If u_d_value <> "" then
		sUDCart = sUDCart & ("<tr><td colspan='6' class='normal'>"&u_d_name)
		if iEditable<>0 then
		    if isNull(User_Defined_Fields_Format) or User_Defined_Fields_Format = "" then
			    sUDCart = sUDCart & ("<BR><textarea name='"&u_d_number&"_"&Line_id&"' cols='35' rows='2'>"&u_d_value&"</textarea>")
		    else
			    sUDCart = sUDCart & ("&nbsp;<input name='"&u_d_number&"_"&Line_id&"' type=text "&User_Defined_Fields_Format&" value='"&u_d_value&"'>")
		    end if
		    sUDCart = sUDCart & ("<input type='hidden' name='"&u_d_number&"_old_"&Line_id&"' value='"&u_d_value&"'>")
	    else
            sUDCart = sUDCart & ("="&u_d_value)
	    end if
	    sUDCart = sUDCart & ("</td></tr>")
	End If
    fn_create_ud=sUDCart
end function
%>