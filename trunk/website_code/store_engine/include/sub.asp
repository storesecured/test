<!--#include file="gift_misc.asp"-->
<!--#include virtual="common/common_functions.asp"-->

<%
function fn_is_credit_card(sPayment_Method)
    sCreditCards="Visa,Mastercard,Discover,American Express,Diners Club"
    if Is_In_Collection(sCreditCards,sPayment_Method,",") then
        fn_is_credit_card=1
    else
        fn_is_credit_card=0
    end if
    fn_print_debug "fn_is_credit_card="&fn_is_credit_card
end function

function fn_is_realtime(sPayment_Method)
    sOtherPayments="eCheck,Debit Card,Checks By Net,Paypal,PayPal-ExpressCheckout"
    if Is_In_Collection(sOtherPayments,sPayment_Method,",") then
        fn_is_realtime=1
    else
        fn_is_realtime=fn_is_credit_card(sPayment_Method)
    end if
    fn_print_debug "fn_is_realtime="&fn_is_realtime
end function

function fn_is_third_party_gateway(sGateway)
    sPostGateways = "8,11,12,17,19,20,22,24"
    if Is_In_Collection(sPostGateways,sGateway,",") then
        fn_is_third_party_gateway=1
    else
        fn_is_third_party_gateway=0
    end if
    fn_print_debug "fn_is_third_party_gateway="&fn_is_third_party_gateway
end function

function fn_is_paypal_express(sPayment_Method,sGateway)
    if sPayment_Method="PayPal-ExpressCheckout" then
        fn_is_paypal_express=1
    else
        fn_is_paypal_express=0    
    end if
end function

function fn_is_paypal(sPayment_Method,sGateway)
	fn_print_debug "gateway="&sGateway&" and sPayment_Method="&sPayment_Method
    if (cint(sGateway) = cint(4) or sPayment_Method  = "Paypal") then
        	fn_print_debug "paypal gateway=yes"
	   fn_is_paypal=1
    else
        fn_is_paypal=0    
    end if
end function

function fn_is_checksbynet(sPayment_Method,sGateway)
    if (sGateway = 33 or sPayment_Method  = "Checks By Net") then
        fn_is_checksbynet=1
    else
        fn_is_checksbynet=0    
    end if
end function

function fn_display_address(iCondensed)
    sAddr=""
    if Store_Company<>"" then
      sAddr=sAddr&("<b>"&Store_Company&"</b><BR>")
    end if
    if iCondensed=1 then
        'sAddr=sAddr&(Site_Name&"<BR>")
	end if
	if iCondensed=1 and (Store_Address1<>"" or Store_Address2<>"") then
        sAddr=sAddr&(Store_address1&"&nbsp;"&Store_address2&"<BR>")
    elseif iCondensed=0 and (Store_Address1<>"" or Store_Address2<>"") then
        if Store_Address1<>"" then
            sAddr=sAddr&(Store_address1&"<BR>")
        end if
        if Store_address2<>"" then
            sAddr=sAddr&(Store_address2&"<BR>")
        end if 
    end if
	if Store_City<>"" then
        sAddr=sAddr&(Store_City&", ") 
    end if
    if Store_State <>"" then
        sAddr=sAddr&(Store_State)
    end if
    if Store_Zip<>"" then
        sAddr=sAddr&("&nbsp;"&Store_Zip)
    end if
    sAddr=sAddr&("<BR>")
    if Store_Country<>"" then
	    sAddr=sAddr&(Store_Country&"<BR>")
	end if
	if Store_Phone<>"" then
        sAddr=sAddr&("Phone: "&Store_Phone)
        if Store_Fax<>"" then
            if iCondensed=1 then
                sAddr=sAddr&(", Fax: "&Store_Fax)
            else
                sAddr=sAddr&("<BR>Fax: "&Store_Fax)
            end if
        end if
        sAddr=sAddr&("<BR>")
    end if
    if iCondensed=1 then
	    sAddr=sAddr&(Store_Email)
	end if
	fn_display_address=sAddr
end function
' ================================================================
sub sub_require_login()
    if (cid = 0) then
        fn_redirect Switch_name&"check_out.asp?ReturnTo="&ReturnTo
    end if
end sub


' ================================================================ 
' LOGOUT A CUSTOMER FROM STORE (UPDATING STORE_SHOPPERS TABLE)
Sub LogOut_Customer()
	sql_update = "exec wsp_customer_logout "&Store_Id&","&Shopper_Id&";"
	sql_update = sql_update & fn_recalc_Cart()
	fn_print_debug sql_update
	conn_store.Execute sql_update

End Sub
' ================================================================
' LOGIN A CUSTOMER TO THE STORE (UPDATING STORE_SHOPPERS TABLE WITH A VALID CID)
Sub sub_login_customer(cid)
    sql_update = "exec wsp_customer_login "&store_id&","&Shopper_Id&",'','',"&cid&";"
	sql_update = sql_update & fn_recalc_Cart()
	fn_print_debug sql_update
	conn_store.Execute sql_update
End Sub

' ================================================================
' LOGIN A CUSTOMER TO THE STORE (UPDATING STORE_SHOPPERS TABLE WITH A VALID CID)
function fn_recalc_Cart()
	sql_cart = "exec wsp_cart_recalc "&Store_Id&","&Shopper_Id&",'"&Groups&"';"
	fn_recalc_Cart = sql_cart
End function
' ================================================================ 
' CHECK IF THE SHOPPING CART IS EMPTY OR NOT
Function fn_payment_message(sPayment)
	fn_payment_message = ""
	sql_select = "select Payment_method_message from store_payment_methods where store_id="&store_id&" and payment_name='"&sPayment&"';" 
	fn_print_debug sql_select
	set rs_Store_payment = server.CreateObject("ADODB.Recordset")
	rs_Store_payment.open sql_select,conn_store,1,1
	if not rs_Store_payment.eof then
		fn_payment_message = rs_Store_payment("payment_method_message")
	end if
	rs_Store_payment.Close
	set rs_Store_payment=Nothing 
End Function 
' ================================================================ 
' CHECK IF THE SHOPPING CART IS EMPTY OR NOT
Function fn_country_code(sCountry)
	fn_country_code = ""
	sql_select = "exec wsp_country_code '"&sCountry&"';"
	fn_print_debug sql_select
	session("sql")=sql_select
	set rs_Store_country = server.CreateObject("ADODB.Recordset")
	rs_Store_country.open sql_select,conn_store,1,1
	if not rs_Store_country.eof then
		fn_country_code = rs_Store_country("country_code")
	end if
	rs_Store_country.Close
	set rs_Store_country=Nothing 
End Function 

' ================================================================ 
' CHECK IF THE SHOPPING CART IS EMPTY OR NOT
Function fn_country_name(sCountryCode)
	fn_country_name = ""
	sql_select = "exec wsp_country_name '"&sCountryCode&"';"
	session("sql")=sql_select
	fn_print_debug sql_select
	set rs_Store_country = server.CreateObject("ADODB.Recordset")
	rs_Store_country.open sql_select,conn_store,1,1
	if not rs_Store_country.eof then
		fn_country_name = rs_Store_country("country")
	end if
	rs_Store_country.Close
	set rs_Store_country=Nothing 
End Function  

' ================================================================ 
' CHECK IF THE SHOPPING CART IS EMPTY OR NOT
Function fn_Is_Cart_Full()
	fn_Is_Cart_Full = 0
	Sql_Cart_Totals = "select is_full from dbo.wf_is_cart_full("&Store_Id&","&Shopper_ID&")"
	fn_print_debug Sql_Cart_Totals
	set rs_Store_cart = server.CreateObject("ADODB.Recordset")
	rs_Store_cart.open Sql_Cart_Totals,conn_store,1,1
	if not rs_Store_cart.eof then
		fn_Is_Cart_Full = rs_Store_cart("is_full")
	end if
	rs_Store_cart.Close
	set rs_Store_cart=Nothing 
End Function  
' ================================================================
' COMPUTE THE SHOPPING CART TOTAL SALE PRICE
Function fn_get_price_total()
	Set rs_storeprice = Server.CreateObject("ADODB.Recordset")
	Sql_Cart_Totals = "exec wsp_cart_total_price "&store_Id&","&Shopper_ID&";"
	fn_print_debug Sql_Cart_Totals
	rs_storeprice.open Sql_Cart_Totals,conn_store,1,1
	Cart_Total = rs_storeprice("Cart_Total")
	rs_storeprice.Close
	set rs_storeprice = nothing
	fn_get_price_total=cart_total
End Function

' ================================================================
' CHECK IF CUSTOMER BELONGS TO A CUSTOMER GROUP
Function fn_Get_Cid_Groups(cid)
    sql_select = "exec wsp_customer_groups "&Store_Id&","&Cid
    fn_print_debug sql_select
    set cidfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select,ciddata,cidfields,noRecords)
    FOR cidrowcounter= 0 TO cidfields("rowcount")
        groups=groups&ciddata(cidfields("group_id"),cidrowcounter)&","
    next

    fn_Get_Cid_Groups=groups
End Function


' ================================================================
' CHECK IF AN ITEM IS IN A COLECTION
Function fn_Is_Instr_Collection(Collection,one_item,Del)
	fn_Is_Instr_Collection = False
	My_Collection_Array = Split(Collection,Del)
	' loop over array to find Collection , we will trim just in case we have spaces ...
	fn_print_debug "Element= "&one_item
	For each Collection_element in My_Collection_Array
		sElement=lcase(trim(CStr(Collection_element)))
		sCheckElement=lcase(CStr(one_item))
  		if sElement<>"" and instr(sCheckElement,sElement)>0 then
			fn_Is_Instr_Collection = True
			fn_print_debug sElement&"=True"
			exit for
		end if
	next
end function

' ================================================================ 
' CHECK IF AN ITEM IS INSIDE A COLLECTION
Function Is_In_Collection(Collection,one_item,Del)
	Is_In_Collection = False

	if not isNull(Collection) and not isNull(one_item) then
	  My_Collection_Array = Split(Collection,Del)
	  For each Collection_element in My_Collection_Array
		  if trim(Cstr(Collection_element)) = Cstr(one_item) then
			  Is_In_Collection = True
			  exit for
		  end if
	  next
	end if
end function


' ================================================================

Function CheckReferer()

strHost = replace(Request.ServerVariables("HTTP_HOST"),"www.","")
strReferer = Request.ServerVariables("HTTP_REFERER")
if strReferer <> "" and instr(strReferer,"http://") = 1 then
  strReferer = Right(strReferer, Len(strReferer) - (InStr(1, strReferer, "://") + 2))
  iLength=InStr(1, strReferer, "/") - 1
  if iLength<0 then
  	iLength=len(strReferer)
  end if
  strReferer = replace(Left(strReferer, iLength),"www.","")
  If strReferer = strHost Then
	  blnCheckReferer = True
  Else
		Host1 = replace(replace(Site_Name_Orig,"www.",""),"http://","")
		Host2 = replace(replace(Store_Domain,"www.",""),"http://","")
		if not(isNull(Store_Domain2)) then
			Host3 = replace(replace(Store_Domain2,"www.",""),"http://","")
	     end if

		if Host1 = strReferer or Host2 = strReferer or Host3 = strReferer then
			blnCheckReferer = True
		else
			blnCheckReferer = False
		end if
  End If
  CheckReferer = blnCheckReferer
else
  CheckReferer = true
end if
End Function

' ================================================================



function unencode(sString) 

sString = replace(sString,"%255F","_")
sString = replace(sString,"%5F","_")
sString = replace(sString,"%252D","-")
sString = replace(sString,"%252F","/")
sString = replace(sString,"%2F","/")
sString = replace(sString,"%252E",".")
sString = replace(sString,"%2E",".")
sString = replace(sString,"%253A",":")
sString = replace(sString,"%3A",":")
sString = replace(sString,"%252C",",")
sString = replace(sString,"%2C",",")
sString = replace(sString,"%20"," ")
sString = replace(sString,"%26","&")
sString = replace(sString,"%3F","?")
sString = replace(sString,"%21","!")
sString = replace(sString,"%2D","-")
sString = replace(sString,"%3D","=")
sString = replace(sString,"%3B",";")


unencode = sString

end function

function fn_purchase_decline(oid,sReason)
    fn_purchase_unlock oid
    fn_redirect Switch_Name&"error.asp?Message_id=101&Message_Add="&Server.UrlEncode(sReason)
end function 

function fn_check_protected (iDepartment_Id,Db_Full_Name)
    if Protected_Page_Access=0 or Protected_Page_Access="" then
        sql_select = "exec wsp_dept_protected "&store_id&","&iDepartment_Id&",'"&Db_Full_Name&"';"
        fn_print_debug sql_select
        rs_store.open sql_select, conn_store, 1, 1
		if not rs_store.eof then
			rs_store.close
			fn_page_protected
		end if
		rs_store.close
    end if
end function

function fn_stripHTML(strHTML)
'Strips the HTML tags from strHTML
  if strHTML <> "" then
	 Dim objRegExp, strOutput
	 Set objRegExp = New Regexp

	 objRegExp.IgnoreCase = True
	 objRegExp.Global = True
	 objRegExp.Pattern = "<(.|\n)+?>"

	 'Replace all HTML tag matches with the empty string
	 strOutput = objRegExp.Replace(strHTML, "")

	 'Replace all < and > with &lt; and &gt;
	 strOutput = Replace(strOutput, "<", "&lt;")
	 strOutput = Replace(strOutput, ">", "&gt;")
	 strOutput = Replace(strOutput, """", "")

	 fn_stripHTML = strOutput	  'Return the value of strOutput
	 Set objRegExp = Nothing
  else
	 fn_stripHTML = strHTML
  end if
end function

function fn_page_protected ()
    sUrl=CurrentFilename
    if Protected_page_access=0 or Protected_page_access="" then
        if request.querystring<>"" then
            sUrl = sUrl&"?"&request.querystring
        end if
        fn_print_debug "page is protected, redirecting"
        fn_redirect Switch_Name&"check_out.asp?Protected=Yes&ReturnTo="&ReturnTo
    end if
end function

function fn_page_group ()
    sUrl=CurrentFilename
    if not(Is_In_Collection(Groups,sCustomer_Group,",")) then
        if request.querystring<>"" then
            sUrl = sUrl&"?"&request.querystring
        end if
        fn_redirect Switch_Name&"check_out.asp?Protected=Yes&ReturnTo="&ReturnTo
    end if
end function

function fn_transform_deptstring (sDeptName)
    'take a querystring dept and transform to a full name
    'fn_print_debug "in transform deptstring "&sDeptName
    sDeptName=fn_replace(sDeptName,"^","_")
    sDeptName=fn_replace(sDeptName,"-","_")
	sDeptName=fn_replace(sDeptName,"~","%")

    fn_transform_deptstring = fn_replace(sDeptName,"/"," > ")

end function

function fn_item_missing (sFullName, sItemName)
	fn_print_debug "in item missing "&sFullName
	'user went to the url for an item that could not be found, try to find a new place for them
	sql_select = "select item_page_name, full_name from sv_items_dept_combine WITH (NOLOCK) where store_id="&Store_Id&" and item_page_name='"&sItemName&"' and visible=1"
	fn_print_debug sql_select
	rs_store.open sql_select,conn_store,1,1
	if not rs_store.eof and not rs_store.bof then
		new_full_name=rs_store("full_name")
	else
	    new_full_name=""
    end if
	rs_store.close
	
	if new_full_name<>"" then
		'we found a match so redirect
		fn_redirect_perm fn_item_url(new_full_name,sItemName)
	else
		fn_redirect_perm fn_dept_url(sFullName,"")
	end if
    
end function

function fn_dept_missing (sFullName)
	fn_print_debug "in dept missing"
	'user went to the url for a dept that could not be found
	if instr(sFullName," > ") > 0 then
		sArrayDept = split(sFullName," > ")
		sLevels = ubound(sArrayDept)
		
		'try to find this department in another heirarchy
		sql_select = "select department_name, full_name from store_dept WITH (NOLOCK) where store_id="&store_id&" and department_name='"&sArrayDept(sLevels)&"'"
		fn_print_debug sql_select
		rs_store.open sql_select,conn_store,1,1
		if not rs_store.eof and not rs_store.bof then
			department_name=rs_store("department_name")
			full_name=rs_store("full_name")
		else
			department_name=""
			full_name=""
		end if
		rs_store.close

		if full_name<>"" then
			'found this department elsewhere so redirect
			fn_redirect_perm fn_dept_url(full_name,"")
		end if

		'didnt find department so try one level up in the dept list
		Do while  i< sLevels
			if sShorterName="" then
				sShorterName = sArrayDept(i)
			else
				sShorterName = sShorterName & " > " & sArrayDept(i)
			end if
			i=i+1
		loop
		fn_print_debug "should redirect to "&sShorterName
		fn_redirect_perm fn_dept_url(sShorterName,"")
	else
		'we are already at a top level dept so go to main browse page
		fn_redirect_perm fn_dept_url("","")
	end if
end function

sub sub_check_quantity ()
     sql_select = "exec wsp_cart_qty_check "&store_id&","&Shopper_id&";"
        fn_print_debug sql_select
        rs_store.open sql_select, conn_store, 1, 1
		if not rs_store.eof then
			item_name=rs_store("item_name")
			quantity=rs_store("quantity")
			rs_store.close
			if quantity<=0 then
				fn_error item_name&" is now out of stock.  Please remove it from your cart before proceeding."
			else
				fn_error "There are only "&quantity&" "&item_name&" in stock at this time.  Please decrease your quantity before proceeding."
			end if
		end if
		rs_store.close
end sub

sub sub_check_minimum ()
	if cdbl(Minimum_Amount)>cdbl(0) then
	    set rs_Storemin =  server.createobject("ADODB.Recordset")
	    sql_select = "exec wsp_cart_total_price "&store_id&","&Shopper_id&";"
	    fn_print_debug sql_select
	    rs_Storemin.open sql_select, conn_store, 1, 1
	    total_Price=rs_Storemin("Cart_Total")
	    rs_Storemin.close
	    set rs_Storemin=Nothing
	
	    if cdbl(total_Price)<cdbl(Minimum_Amount) then
	    	fn_error "You must order at least $" & Minimum_Amount
	    end if
    end if

end sub

function fn_old_dept_url (iCateg_id)
	if fn_isID(iCateg_id)=0 then
		'invalid id, redirect to homepage
		fn_redirect_perm Switch_Name
	end if
         'person is coming from old url, send them to the new url right away
            sql_select_items="wsp_dept_transform "&store_id&","&iCateg_id
            fn_print_debug sql_select_items
            session("sql")=sql_select_items
	        rs_store.open sql_select_items,conn_store,1,1
	        if not rs_store.eof then
		        full_name=rs_store("full_name")
	        else
		        full_name=""
	        end if
	        rs_store.close
	        sUrl = fn_dept_url(full_name,"")
	        fn_print_debug "url="&sUrl
	        if request.queryString<>"" then
	           fn_print_debug "in querystring ="&request.queryString
                   sQuerystring = ""

	           for each key in request.queryString
	               fn_print_debug "key="&key
	               key=lcase(key)
		       if key<>"name" and key<>"categ_id" and key<>"item_id" and key<>"parent_ids" and key<>"page_id" and key<>"url_string" and key<>"jump_to_page" and key<>"sub_department_id" then
		          if sQuerystring="" then
                             sQuerystring=key&"="&request.queryString(key)
                          else
                             sQuerystring=sQuerystring&"&"&key&"="&request.queryString(key) 
                          end if
		       end if
                   next
                   if sQuerystring<>"" then
	              sUrl=sUrl&"?"&sQuerystring
	           end if
	        end if
	        if store_id<>"10506" then
	            fn_redirect_perm sUrl
	        else
	            q_Dept_Page_Name = fn_replace(full_name," > ","/")
            end if
            fn_old_dept_url = q_Dept_Page_Name
end function

function fn_old_item_url (categ_id,Item_ID)
                    if (not isNumeric(Item_ID) or Item_ID = "") or (not isNumeric(categ_id)) then
		            full_name=""
		            fn_redirect_perm fn_dept_url(full_name,"")
	            else
                    sql_select_items="exec wsp_item_transform "&store_id&","&Item_ID&";"
                    rs_store.open sql_select_items,conn_store,1,1
                    if not rs_store.eof and not rs_store.bof then
                        item_page_name=rs_store("item_page_name")
                        full_name=rs_store("full_name")
                    elseif isNumeric(categ_id) and categ_id<>"" then
                        rs_store.close
				    sql_select_items="exec wsp_dept_transform "&store_id&","&categ_id&";"
	                   fn_print_debug sql_select_items
				     rs_store.open sql_select_items,conn_store,1,1
	                    if not rs_store.eof and not rs_store.bof then
		                    full_name=rs_store("full_name")
	                    else
	                        full_name=""
                        end if
                        rs_store.close
                        fn_redirect_perm fn_dept_url(full_name,"")
                    else
                    	fn_redirect_perm Switch_Name
				end if
                    rs_store.close
                    sUrl = fn_item_url(full_name,item_page_name)
                    if request.queryString<>"" then
	           fn_print_debug "in querystring ="&request.queryString
                   sQuerystring = ""

	           for each key in request.queryString
	               fn_print_debug "key="&key
	               key=lcase(key)
		       if key<>"name" and key<>"categ_id" and key<>"item_id" and key<>"parent_ids" and key<>"page_id" and key<>"url_string" and key<>"jump_to_page" and key<>"sub_department_id" then
		          if sQuerystring="" then
                             sQuerystring=key&"="&request.queryString(key)
                          else
                             sQuerystring=sQuerystring&"&"&key&"="&request.queryString(key)
                          end if
		       end if
                   next
                   if sQuerystring<>"" then
	              sUrl=sUrl&"?"&sQuerystring
	           end if
	        end if
                    fn_redirect_perm sUrl
	            end if
end function

function fn_old_page_url (iPage_Id)
    sLink = Switch_Name&"page_"&iPage_Id&".asp"
    fn_old_page_url = sLink
    fn_redirect sLink
end function

function fn_create_image (sImagePath,sAltText)
    if isNull(sImagePath) or sImagePath = "" or len(sImagePath)<1 Then
	    fn_create_image = "&nbsp;"
	else
		sAlt=" alt='"&checkstringforQ(sAltText)&"'"
	    if Instr(sImagePath,"http://") > 0 then
		    fn_create_image = "<IMG Src='"&sImagePath&"'"&sAlt&" border='0'>"
	    else

		    fn_create_image = "<IMG Src='"&Switch_Name&replace("images/"&sImagePath,"//","/")&"'"&sAlt&" border='0'>"
	    end if
    end if
end function

function fn_isID(iId)
	if isNumeric(iId) and instr(iId,",")=0 and instr(iId,".")=0 then
		fn_isID=1
	else
		fn_isID=0
	end if
end function

sub sub_write_log (sDebug)
	if 1=0 then
		sIP=Request.ServerVariables("REMOTE_ADDR")
		sMessageLog = replace(sDebug,"'","''")
		if store_id="" then
			this_Store_id=0
		else
			this_Store_id=store_id
		end if
		sql_insert = "insert into store_google_log (store_id,client_ip,debug_message) values ("&this_Store_id&",'"&sIP&"','"&sMessageLog&"');"
	    	    session("sql")=sql_insert
		    conn_store.execute sql_insert
    end if
end sub

%>
