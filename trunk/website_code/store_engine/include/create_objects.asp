<%

function fn_create_links_object (link_page,link_text) 
	if fn_create_loginlogout(link_text)=true then  
        link_text = Replace(link_text," ",chr(160))
        if link_text = "" then
            sLinkText = "<a href='"&Site_Name&link_page&"' class='link'>"
        else
            sLinkText = "<a href='"&Site_Name&link_page&"' class='link'>"&link_text&"</a>"
        end if
    else
        sLinkText=""
	end if
	fn_create_links_object=sLinkText
end function

function fn_create_instance_date()
	fn_create_instance_date = "<font class='normal'> "&Replace(FormatDateTime(Now(),1)," ",chr(160))&"</font>"
end function

function fn_create_instance_Hello_user()
    sHelloText = "<tr><td width='100%' class='normal'>"
    if Cid = 0 then
    else
		     sHelloText = sHelloText & "Hello, "&Replace(first_name," ",chr(160))&"&nbsp;"&Replace(last_name," ",chr(160))
    end if
    sHelloText = sHelloText & "</td></tr>"
    fn_create_instance_Hello_user=sHelloText
end function

function fn_create_instance_Links_horizontal()
    fn_create_instance_Links_horizontal = fn_create_instance_links(" : ")
end function

function fn_create_instance_links(sStart,sSeparator,sEnd)
    sLinks = sStart
    sLinks = fn_create_links_object("store.asp","Home")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("search.asp","Search")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("Browse_dept_items.asp","Browse")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("Check_Out.asp","Login")
	sLinks = sLinks & fn_create_links_object("log_out.asp","Logout")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("my_account.asp","My Account")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("show_big_cart.asp","View Cart")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("Before_Payment.asp","Check Out")
	sLinks = sLinks & sSeparator
	sLinks = sLinks & fn_create_links_object("aboutus.asp","About Us")
	sLinks = sLinks & sEnd
	fn_create_instance_links = sLinks
end function

function fn_create_instance_Links_vertical()
    fn_create_instance_Links_vertical = fn_create_instance_links("<tr><td>","</td></tr><tr><td>","</td></tr>")
end function

function fn_create_instance_newsletter_box
	sNewsletterText = "<form action='"&Site_Name&"signup_action.asp' method='post'><input type='Text' name='Email_Address' size='15'><INPUT type='hidden'  name=Email_Address_C value='Re|Email|0|50|||Email Address'><BR>"
	sNewsletterText = sNewsletterText & fn_create_action_button ("Button_image_Continue", "Continue", "Continue")
	sNewsletterText = sNewsletterText & "</form>"
	fn_create_instance_newsletter_box=sNewsletterText
end function

function fn_create_instance_Search_box
	sSearchText = "<form action='"&Site_Name&"search_items.asp' method='post'><input type='Text' name='Search_Text' size='15'><BR>"
	sSearchText = sSearchText & fn_create_action_button ("Button_image_Search", "Search_Store", "Search")
	sSearchText = sSearchText & "</form>"
	fn_create_instance_Search_box=sSearchText
end function

function fn_create_instance_Small_cart(useTline)
	sql_Show_big_cart = "exec wsp_cart_display "&store_id&","&Shopper_ID&";"
	fn_print_debug sql_Show_big_cart
	set myfieldscart=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_Show_big_cart,mydatacart,myfieldscart,noRecordscart)

	big_text_size = text_size +1
	sCartText = "<table cellspacing='0' cellpadding='2' border='0' width='100%'>"
		If useTLine then
			sCartText = sCartText & "<tr><td colspan='3' class='normal'><STRONG>My Shopping Cart</STRONG></td></tr>"
		End If
		sCartText = sCartText & "<tr><td class='small'><STRONG>Qty</STRONG></td>"&_
		    "<td class='small'><STRONG>Item</STRONG></td>"&_
		    "<td class='small'><STRONG>Total</STRONG></td></tr>"
 
		Line_id = 1
		sTotal=0
		if noRecordscart = 0 then
		FOR rowcountercart= 0 TO myfieldscart("rowcount")

			'calculate ext price for item ...
			Quantity = mydatacart(myfieldscart("quantity"),rowcountercart)
			Sale_Price = mydatacart(myfieldscart("sale_price"),rowcountercart)
			Ext_Sale_Price = Quantity*sale_price
			sTotal=sTotal+Ext_Sale_Price

			if Ext_Sale_Price = 0 then
				Ext_Sale_Price = "<i>Free</i>"
			else
				Ext_Sale_Price = Currency_Format_Function(Ext_Sale_Price)
			End if
			if Sale_Price = 0 then
				Sale_Price = "<i>Free</i>"
			else
				Sale_Price = Currency_Format_Function(sale_price)
			End if
               Full_name = mydatacart(myfieldscart("full_name"),rowcountercart)

			Item_Id = mydatacart(myfieldscart("item_id"),rowcountercart)
			Item_Page_Name = mydatacart(myfieldscart("item_page_name"),rowcountercart)
			Item_Name = mydatacart(myfieldscart("item_name"),rowcountercart)
			if len(Item_Name)>15 then
				Item_name = left(mydatacart(myfieldscart("item_name"),rowcountercart),15)&"..."
			else
				Item_name = mydatacart(myfieldscart("item_name"),rowcountercart)
			end if
			Custom_Link=mydatacart(myfieldscart("custom_link"),rowcountercart)

			sCartText = sCartText & ("<tr><td class='small'>"&Quantity&"</td>")
			
			if Custom_Link <> "" then
			    sCartLink=Custom_Link
			else
			    sCartLink=fn_item_url(full_name,Item_Page_Name)
			end if
			sCartText = sCartText & ("<td><a href='"&sCartLink&"' class='link'><font size=-2>"&Item_name&"</font></a></td>"&_
			    "<td class='small'>"&Ext_Sale_Price&"</td></tr>")
			Line_id = Line_id + 1
		Next
		end if
		if useTline then
		  sCartText = sCartText & ("<tr><td colspan=2></td><td class='small'><STRONG>"&Currency_Format_Function(sTotal)&"</STRONG></td></tr>")
		end if
	sCartText = sCartText & "</table>"
	
	set myfieldscart = Nothing
	fn_create_instance_Small_cart=sCartText
end function

function fn_create_instance_login()
	
    sLoginText = "<!-- login object starts --><form method='POST' action='"&Site_Name&"Check_Out_Action.asp'"
    if Cid = 0 then
	    sLoginText = sLoginText & ("><input type='Hidden' name='Form_Name' value='Check_Out'><table><tr>"&_
	        "<td width='24%' height='27' class='small'>Login</td>"&_
	        "<td width='76%' height='27' colspan='2'>"&_
	        "<input type='text' name='User_id' size='22'>"&_
	        "<input type='hidden' name='User_id_C' value='Re|String||||' ></td></tr><tr>"&_
	        "<td width='24%' height='25' class='small'>Password</td>"&_
	        "<td width='76%' height='25' colspan='2'>"&_
	        "<input type='password' name='Password' size='22'>"&_
	        "<input type='hidden' name='Password_C' value='Re|String|||' ></td>"&_
	        "</tr><tr><td width='24%' height='27'></td>"&_
	        "<td width='38%' height='27'>"&_
	        fn_create_action_button ("Button_image_Login", "Login", "Login")&_
            "</td><td width='38%' height='27'></td></tr></table>")

    else
		sLoginText = sLoginText & (" name=log_out_yes><input type='Hidden' name='Form_Name' value='Log_Out_Yes'>"&_
		    "<tr><td>" & fn_create_action_button ("Button_image_Logout", "Logout", "Logout") )
	end if
    sLoginText = sLoginText & "</form><!-- login ends -->" 

	fn_create_instance_login = sLoginText
end function

function fn_create_instance_Select_box_Depts

    sql_select = "exec wsp_dept_select_box "&Store_Id&";"
    set myfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

    if noRecords = 0 then
     sDeptText = chr(13) & "<SCRIPT LANGUAGE='JavaScript'>" & chr(13) & "function surfto(form) {" & chr(13) & "var myindex=form.Department_ID.selectedIndex " & chr(13) & "if (form.Department_ID.options[myindex].value != '0') { " & chr(13) & "location=form.Department_ID.options[myindex].value;}}" & chr(13) & "</SCRIPT><table><tr><td>" & "<form action='"&Site_Name&"items/list.htm' method='POST' >" & "<select name='Department_ID' onChange='surfto(this.form)'>" & "<option selected value='0'>Select Dept</option>'" & "<option value='0'>-----------</option>"
     FOR rowcounter= 0 TO myfields("rowcount")
	    Department_name = replace(mydata(myfields("department_name"),rowcounter),"'","&prime;")
	    sFullName=mydata(myfields("full_name"),rowcounter)
	    sDeptText = sDeptText & ("<option value='"&fn_dept_url(sFullName,"")&"'>"&Department_name&"</option> ")
     Next
     sDeptText = sDeptText & "</select></form></td></tr></table>"
    end if

    set myfields = Nothing
    
    fn_create_instance_Select_box_Depts = sDeptText

end function

function fn_putCustomHTML(theHTML)
	'fn_print_debug theHTML
	'fn_print_debug sPage_top
    theHTML = fn_replace(theHTML,"%OBJ_PAGE_TOP_OBJ%",sPage_top)
    theHTML = fn_replace(theHTML,"%OBJ_PAGE_BOTTOM_OBJ%",sPage_bottom)
    theHTML = fn_replace(theHTML,"%OBJ_PAGE_FORM_OBJ%",sPage_form_content)
    theHTML = fn_replace(theHTML,"OBJ_HEAD_OBJ",sPage_Head)
    theHTML = fn_replace(theHTML,"OBJ_TITLE_OBJ",sMeta_Title)
    theHTML = fn_replace(theHTML,"OBJ_KEYWORD_OBJ",sMeta_Keywords)
    theHTML = fn_replace(theHTML,"OBJ_DESCRIPTION_OBJ",sMeta_Description)
    theHTML = fn_replace(theHTML,"URL_STRING",url_string)
    theHTML = fn_replace(theHTML,"OBJ_SWITCH_NAME",Switch_Name)
    theHTML = fn_replace(theHTML,"OBJ_TITLE_OBJ",fn_get_querystring("Name"))
    theHTML = fn_replace(theHTML,"%OBJ_BUDGET_OBJ%",budget_left)


		if instr(theHTML,"%OBJ_NAVBUTTON_") then
		    fn_print_debug "in navbutton"
	        sLoop=fn_make_nav("button",NavButton_1_Loop,1)    
	        theHTML = Replace(theHTML,"%OBJ_NAVBUTTON_1_OBJ%",sLoop)
	    end if
        
	    if inStr(theHTML,"%OBJ_AFFILIATE_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_AFFILIATE_OBJ%",Came_From)
	    end if
	    if inStr(theHTML,"%OBJ_PRICE_PLAIN_OBJ%") then
	      theHTML = Replace(theHTML,"%OBJ_PRICE_PLAIN_OBJ%",fn_Get_Price_Total())
	    end if
	    if inStr(theHTML,"%OBJ_PRICE_OBJ%") then
		 fn_print_debug "fn_Get_Price_Total()="&fn_Get_Price_Total()
	      theHTML = Replace(theHTML,"%OBJ_PRICE_OBJ%",Currency_Format_Function(fn_Get_Price_Total()))
	    end if
	    if inStr(theHTML,"%OBJ_ITEMS_OBJ%") then
	      theHTML = Replace(theHTML,"%OBJ_ITEMS_OBJ%",fn_get_items_total())
	    end if
	    if instr(theHTML,"%OBJ_DATE_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_DATE_OBJ%",fn_create_instance_date)
	    end if
	    if instr(theHTML,"%OBJ_HELLO_USER_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_HELLO_USER_OBJ%",fn_create_instance_Hello_user)
	    end if
	    if instr(theHTML,"%OBJ_LINKS_HORIZONTAL_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_LINKS_HORIZONTAL_OBJ%",fn_create_instance_Links_horizontal)
	    end if
	    if instr(theHTML,"%OBJ_LINKS_VERTICAL_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_LINKS_VERTICAL_OBJ%",fn_create_instance_Links_vertical)
	    end if

	    if instr(theHTML,"%OBJ_SEARCH_BOX_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_SEARCH_BOX_OBJ%",fn_create_instance_Search_box)
	    end if
	    if instr(theHTML,"%OBJ_SMALL_CART_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_SMALL_CART_OBJ%",fn_create_instance_Small_cart(true))
	    end if
	    if instr(theHTML,"%OBJ_SMALL_CARTNT_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_SMALL_CARTNT_OBJ%",fn_create_instance_Small_cart(false))
	    end if
	    if instr(theHTML,"%OBJ_LOGIN_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_LOGIN_OBJ%",fn_create_instance_login)
	    end if
	    if instr(theHTML,"%OBJ_SELECT_BOX_DEPTS_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_SELECT_BOX_DEPTS_OBJ%",fn_create_instance_Select_box_Depts)
	    end if
	    if instr(theHTML,"%OBJ_NEWSLETTER_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_NEWSLETTER_OBJ%",fn_create_instance_newsletter_box)
	    end if
	    if instr(theHTML,"%OBJ_BANNER_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_BANNER_OBJ%",fn_create_instance_banner())
	    end if
	    if instr(theHTML,"%OBJ_FIRSTNAME_OBJ%") then
	        theHTML = Replace(theHTML,"%OBJ_FIRSTNAME_OBJ%",First_Name)
	    end if
	    
	    if instr(theHTML,"%OBJ_NAV_") then
	        theHTML = Replace(theHTML,"%OBJ_NAV_BUTTONS_OBJ%",Template_Buttons)
		    theHTML = Replace(theHTML,"%OBJ_NAV_LINKS_OBJ%",Template_Links)
	    end if
	    if instr(theHTML,"%OBJ_BUTTON_") then
	        theHTML = fn_replace_button(theHTML)
	    end if
	    if instr(theHTML,"%OBJ_HTTP_") then
	        theHTML = fn_replace_http(theHTML)
	    end if
	    if instr(theHTML,"%OBJ_") and instr(theHTML,"_OBJ%") then
	        theHTML = fn_replace_links(theHTML)
	    end if
	    theHTML = fn_replace(theHTML,"%OBJ_SHOPPER_OBJ%",Shopper_Id)
         theHTML = fn_replace(theHTML,"%OBJ_SYS_IMAGES_OBJ%","images/images_themes/")

		theHTML = fn_replace(theHTML,"%OBJ_STORE_IMAGES_OBJ%","images")
		theHTML = fn_replace(theHTML,"%OBJ_PHONE_OBJ%",Store_Phone)
		theHTML = fn_replace(theHTML,"%OBJ_NAME_OBJ%",Store_Name)
		if instr(theHTML,"%OBJ_TOTAL_OBJ") then
			theHTML = fn_replace(theHTML,"%OBJ_TOTAL_OBJ%",fn_get_grandtotal())
     	end if
     	if instr(theHTML,"%OBJ_SUBTOTAL_OBJ") then
			theHTML = fn_replace(theHTML,"%OBJ_SUBTOTAL_OBJ%",fn_get_total())
     	end if
		theHTML = fn_replace(theHTML,"%OBJ_ORDER_OBJ%",oid)


	fn_putCustomHTML = theHTML
End function

function fn_get_grandtotal ()
	fn_get_grandtotal = 0
	if oid<>0 then
		sql_totals = "select sum(grand_total) as grand_total from store_purchases WITH (NOLOCK) where store_id="&Store_Id&" and (oid="&oid&" or masteroid="&oid&");"
		fn_print_debug sql_totals
		set rs_store_total = conn_store.execute(sql_totals)
		fn_get_grandtotal = rs_store_total("grand_total")
    		set rs_store_total=nothing
	end if
end function

function fn_get_total ()
	fn_get_total = 0
	if oid<>0 then
		sql_totals = "select sum(total) as total from store_purchases WITH (NOLOCK) where store_id="&Store_Id&" and (oid="&oid&" or masteroid="&oid&");"
		fn_print_debug sql_totals
		set rs_store_total = conn_store.execute(sql_totals)
		fn_get_total = rs_store_total("total")
    		set rs_store_total=nothing
	end if
end function

function fn_replace_button(sHtmlText)
    sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_ABOUT_OBJ%",fn_create_instance_button("about"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_LOGIN_OBJ%",fn_create_instance_button("login")&fn_create_instance_button("logout"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_HOME_OBJ%",fn_create_instance_button("home"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_SEARCH_OBJ%",fn_create_instance_button("search"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_BROWSE_OBJ%",fn_create_instance_button("browse"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_MYACCOUNT_OBJ%",fn_create_instance_button("myaccount"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_VIEWCART_OBJ%",fn_create_instance_button("viewcart"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_REGISTER_OBJ%",fn_create_instance_button("register"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_PASTORDERS_OBJ%",fn_create_instance_button("pastorders"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_HELP_OBJ%",fn_create_instance_button("help"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_CHECKOUT_OBJ%",fn_create_instance_button("checkout"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_CONTACTUS_OBJ%",fn_create_instance_button("contactus"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_PRIVACY_OBJ%",fn_create_instance_button("privacy"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_AFFILIATES_OBJ%",fn_create_instance_button("affiliates"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_RETURN_OBJ%",fn_create_instance_button("return"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BUTTON_LOGOUT_OBJ%",fn_create_instance_button("login")&fn_create_instance_button("logout"))

    fn_replace_button=sHtmlText
end function

function fn_replace_http(sHtmlText)
    sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_ABOUT_OBJ%",fn_create_instance_http("about"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_LOGIN_OBJ%",fn_create_instance_http("login")&fn_create_instance_http("logout"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_HOME_OBJ%",fn_create_instance_http("home"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_SEARCH_OBJ%",fn_create_instance_http("search"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_BROWSE_OBJ%",fn_create_instance_http("browse"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_MYACCOUNT_OBJ%",fn_create_instance_http("myaccount"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_VIEWCART_OBJ%",fn_create_instance_http("viewcart"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_REGISTER_OBJ%",fn_create_instance_http("register"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_PASTORDERS_OBJ%",fn_create_instance_http("pastorders"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_HELP_OBJ%",fn_create_instance_http("help"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_CHECKOUT_OBJ%",fn_create_instance_http("checkout"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_CONTACTUS_OBJ%",fn_create_instance_http("contactus"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_PRIVACY_OBJ%",fn_create_instance_http("privacy"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_AFFILIATES_OBJ%",fn_create_instance_http("affiliates"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_RETURN_OBJ%",fn_create_instance_http("return"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HTTP_LOGOUT_OBJ%",fn_create_instance_http("login")&fn_create_instance_http("logout"))

    fn_replace_http=sHtmlText

end function

function fn_replace_links(sHtmlText)
    sHtmlText = fn_replace(sHtmlText,"%OBJ_ABOUT_OBJ%",fn_create_links_object("aboutus.asp","About Us"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_LOGIN_OBJ%",fn_create_links_object("check_out.asp","Login")&fn_create_links_object("log_out.asp","Logout"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HOME_OBJ%",fn_create_links_object("store.asp","Home"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_SEARCH_OBJ%",fn_create_links_object("search.asp","Search"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_BROWSE_OBJ%",fn_create_links_object("Browse_dept_items.asp","Browse"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_MYACCOUNT_OBJ%",fn_create_links_object("my_account.asp","My Account"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_VIEWCART_OBJ%",fn_create_links_object("show_big_cart.asp","View Cart"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_REGISTER_OBJ%",fn_create_links_object("register.asp","Register"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_PASTORDERS_OBJ%",fn_create_links_object("past_orders.asp","Past Orders"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_HELP_OBJ%",fn_create_links_object("store.asp","Help"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_CHECKOUT_OBJ%",fn_create_links_object("before_payment.asp","Checkout"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_CONTACTUS_OBJ%",fn_create_links_object("contactus.asp","Contact Us"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_PRIVACY_OBJ%",fn_create_links_object("privacy.asp","Privacy"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_AFFILIATES_OBJ%",fn_create_links_object("affiliate_program.asp","Affiliates"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_RETURNS_OBJ%",fn_create_links_object("returns.asp","Returns"))
	sHtmlText = fn_replace(sHtmlText,"%OBJ_LOGOUT_OBJ%",fn_create_links_object("check_out.asp","Login")&fn_create_links_object("log_out.asp","Logout"))

    fn_replace_links=sHtmlText
end function

function fn_create_instance_banner()

	sql_banner = "exec wsp_banner_display "&store_id&";"
	rs_store.open sql_banner, conn_store, 1, 1
	if not rs_store.eof then
		sel_id = rs_store("Bann_ID")
		sel_img = rs_store("Image_Path")
	else
	    sel_id=0
	end if
	rs_store.close

    if sel_id=0 then
        fn_create_instance_banner=""
    else
        if instr(sel_img,"http://") then
                sImg=sel_img
        else
                sImg=Switch_Name&"images/"&sel_img
        end if
	    fn_create_instance_banner = "<a href='"&Site_Name&"include/banner_click.asp/BID/"&sel_id&"' target=_blank><img src='"&sImg&"' border='0'></a>"
	    sql_insert = "exec wsp_banner_imp_insert "&store_id&", "&sel_id&", '"&Request.ServerVariables("REMOTE_ADDR")&"';"
	    conn_store.execute sql_insert
	end if

end function

function fn_create_loginlogout (iType)
    if (iType = "login" and Cid > 0) or (iType = "logout" and Cid = 0) then
	    fn_create_loginlogout = false
	else
	    fn_create_loginlogout = true
	end if
end function

function fn_create_instance_button(iType)

    if fn_create_loginlogout(iType)=true then
        if iType = "home" then
	        sURL = Site_Name&"store.asp"
        elseif iType="search" then
	        sURL=Site_Name&"search.asp"
        elseif iType="browse" then
	        sURL = Site_Name&"Browse_dept_items.asp"
        elseif iType="login" then
	        sURL = Site_Name&"Check_Out.asp"
        elseif iType="myaccount" then
	        sURL = Site_Name&"my_account.asp"
        elseif iType="pastorders" then
	        sURL = Site_Name&"Past_orders.asp"
        elseif iType="viewcart" then
	        sURL =Site_Name&"show_big_cart.asp"
        elseif iType="register" then
	        sImagePath = sImage & "register.gif"
        elseif iType="help" then
	        sURL = Site_Name&"store.asp"
        elseif iType="logout" then
	        sURL = Site_Name&"log_out.asp"
        elseif iType="checkout" then
	        sURL = Site_Name&"Before_Payment.asp"
        elseif iType="affiliates" then
	        sURL = Site_Name&"affiliate_program.asp"
        elseif iType="returns" then
	        sURL = Site_Name&"returns.asp"
        elseif iType="privacy" then
	        sURL = Site_Name&"privacy.asp"
        elseif iType="contactus" then
	        sURL = Site_Name&"contactus.asp"
        elseif iType="about" then
	        sURL = Site_Name&"aboutus.asp"
	        iType = "aboutus"
        end if
        sImage = Switch_Name&"images/images_themes/"&Store_Theme&"/"
        sImagePath = sImage & iType & ".gif"

        sButtonText = "<a href="&sURL&">"
        sButtonText = sButtonText & "<img border=0 src="&sImagePath&"></a>"
        fn_create_instance_button = sButtonText
    else
        fn_create_instance_button = ""
    end if
end function

function fn_create_instance_http(iType)

    if fn_create_loginlogout(iType)=true then
        if iType = "home" then
	        sURL = Site_Name&"store.asp"
        elseif iType="search" then
	        sURL=Site_Name&"search.asp"
        elseif iType="browse" then
	        sURL = Site_Name&"Browse_dept_items.asp"
        elseif iType="login" then
	        sURL = Site_Name&"Check_Out.asp"
        elseif iType="myaccount" then
	        sURL = Site_Name&"my_account.asp"
        elseif iType="pastorders" then
	        sURL = Site_Name&"Past_orders.asp"
        elseif iType="viewcart" then
	        sURL =Site_Name&"show_big_cart.asp"
        elseif iType="register" then
            sURL =Site_Name&"register.asp"
        elseif iType="logout" then
	        sURL = Site_Name&"log_out.asp"
        elseif iType="checkout" then
	        sURL = Site_Name&"Before_Payment.asp"
        elseif iType="affiliates" then
	        sURL = Site_Name&"affiliate_program.asp"
        elseif iType="returns" then
	        sURL = Site_Name&"returns.asp"
        elseif iType="privacy" then
	        sURL = Site_Name&"privacy.asp"
        elseif iType="contactus" then
	        sURL = Site_Name&"contactus.asp"
        elseif iType="about" then
	        sURL = Site_Name&"aboutus.asp"
        end if

        sHttpText = "<a href="&sURL&">"
        fn_create_instance_http = sHttpText
    else
        fn_create_instance_http = ""
    end if
end function

' ================================================================
function fn_create_action_button (sButtonType, sName, sText)
    Select Case sButtonType
    case "Button_image_Continue"
        sButton=Button_image_Continue
    case "Button_image_Update"
        sButton=Button_image_Update
    case "Button_image_Cancel"
        sButton=Button_image_Cancel
    case "Button_image_SaveCart"
        sButton=Button_image_SaveCart
    case "Button_image_RetrieveCart"
        sButton=Button_image_RetrieveCart
    case "Button_image_Search"
        sButton=Button_image_Search
    case "Button_image_Delete"
        sButton=Button_image_Delete
    case "Button_image_View"
        sButton=Button_image_View
    case "Button_image_Login"
        sButton=Button_image_Login
    case "Button_image_No"
        sButton=Button_image_No
    case "Button_image_Yes"
        sButton=Button_image_Yes
    case "Button_image_UpdateCart"
        sButton=Button_image_UpdateCart
    case "Button_image_Checkout"
        sButton=Button_image_Checkout
    case "Button_image_Order"
        sButton=Button_image_Order
    case "Button_image_CancelOrder"
        sButton=Button_image_CancelOrder
    case "Button_image_Reset"
        sButton=Button_image_Reset
    case "Button_image_Next"
        sButton=Button_image_Next
    case "Button_image_Prev"
        sButton=Button_image_Prev
    case "Button_image_Up"
        sButton=Button_image_Up
    case "Button_image_ContinueShopping"
        sButton=Button_image_ContinueShopping
    case "Button_image_ProcessOrder"
        sButton=Button_image_ProcessOrder
    end select

    if inStr(sButton,".") > 0 then
        fn_create_action_button = "<input type='Image' src='"&Switch_Name&"images/"&sButton&"' name='"&sName&"'>"
    elseif sButton <> "" then
        fn_create_action_button = "<input type='submit' value='"&sButton&"' name='"&sName&"'>"
    else
        fn_create_action_button = "<input type='submit' value='"&sText&"' name='"&sName&"'>"
    end if
    
end function

Function fn_get_items_total()
	fn_get_items_total = 0
	sql_totals = "exec wsp_cart_total_qty "&Store_Id&","&Shopper_ID&";"
	fn_print_debug sql_totals
	set rs_store_total = conn_store.execute(sql_totals)
	fn_get_items_total = rs_store_total("Qty_Total")
    set rs_store_total=nothing
End Function

function fn_make_nav (sType,sNavTemplate,sMenu_num)
    fn_print_debug "in fn_make_nav"
    
    links = ""
    
    if sNavTemplate<>"" then
        sql_select = "wsp_design_menu "&store_id&","&sMenu_num&",'"&stype&"',"&cint(Session("No_Ecommerce"))&";"
        fn_print_debug sql_select
        set myfields=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

        if noRecords = 0 then
            FOR rowcounter= 0 TO myfields("rowcount")
                fn_print_debug "in loop"
                bCreate = 1
                File_name = mydata(myfields("file_name"),rowcounter)
                Page_name = mydata(myfields("page_name"),rowcounter)
                is_link = mydata(myfields("is_link"),rowcounter)
                menu_type = mydata(myfields("menu_type"),rowcounter)
    		
    		    if menu_type="page" then
                    if File_name = "affiliate_program.asp" and Enable_affiliates = 0 then
	                    bCreate=0
                    elseif File_name = "browse_dept_items.asp" then
	                    File_name="items/list.htm"
                    end if

                    File_Name = fn_page_url (File_Name,is_link)
                else
                    File_Name = fn_dept_url(File_Name,"")
                end if
                if bCreate then
	                numLinks = numLinks + 1
	                sLinkText = "href='"&File_name&"'>"&Page_name
	                links = links & Replace(sNavTemplate,"%OBJ_LINK_OBJ%",sLinkText)
			    end if
            Next
        End if
    end if
    
    fn_make_nav=links

end function
%>
