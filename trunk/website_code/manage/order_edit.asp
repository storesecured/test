<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/country_list.asp"-->
<%

sInstructions="Edits done from the order details page are manual edits only.  Please note that they do not affect inventory quantities, and do not change the amount charged to your customer, etc.  The grand total will be recalculated based the options you have selected."

function getSalePrice(iid)
	sql_select_item = "select * from store_items where store_id = "&store_id&" and item_id = "&iid
	rs_Store.open sql_select_item,conn_store,1,1
	If  rs_store("Retail_Price_special_Discount")>0 AND rs_store("Special_start_date") <= Now() AND rs_store("Special_end_date") >= Now() Then
		Special_Price = rs_store("Retail_Price")*(100-rs_store("Retail_Price_special_Discount"))/100
		getSalePrice=special_price
	else
		getSalePrice=rs_store("retail_price")
	end if
	rs_Store.close
end function

function checkQuantity(quant,iid)
	reterr=""
	if Not IsNumeric(quant) then
		reterr = reterr&"The quantity must be numeric"&"<BR>"
		checkQuantity = reterr
		exit function
	end if
	sql_item = "select Fractional, Quantity_in_stock,Quantity_Control_Number,Quantity_Control, Quantity_Minimum from store_items where Store_id = "&Store_id&" and Item_Id = "&iid
	rs_store.open sql_item,conn_store,1,1
	Quantity_in_stock = Rs_store("Quantity_in_stock")
	Quantity_Control = Rs_store("Quantity_Control")
	Quantity_Control_Number = Rs_store("Quantity_Control_Number")
	Quantity_Minimum = Rs_store("Quantity_Minimum")
	Fractional = cint(Rs_store("Fractional"))
	if IsNull(Quantity_Minimum) then
		Quantity_Minimum = 0
	end if
	rs_store.close
	if Quantity_Control = -1 AND (Quantity_in_stock - quant	< Quantity_Control_Number) then
		reterr = reterr&"You have attempted to order more items then there are in stock"&"<BR>"
		checkQuantity = reterr
		exit function
	end if 
	if cint(quant) < cint(Quantity_Minimum) then
		reterr = reterr&"Quantity ordered must be greater then "&Quantity_Minimum&"<BR>"
		checkQuantity = reterr
		exit function
	end if
	if (Fractional<>-1) then
		if cdbl(round(quant)) <> cdbl(quant) then
			reterr = reterr&"This item cannot be purchased fractionally."&"<BR>"
			checkQuantity = reterr
			exit function
		end if
	end if
	checkQuantity = reterr
end function

Oid = Request.QueryString("Id")
if Oid="" then
   Oid = Request.QueryString("Oid")
end if
if Oid="" or not isNumeric(oid) then
   response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("You must choose a valid order to edit.<BR><BR><a href=orders.asp>Click here to return to the order list.</a>")
end if

'LOAD VALUES FROM STORE_PURCHASES TABLE
sql_select_purchases = "Select cid from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id
rs_Store.open sql_select_purchases,conn_store,1,1
	Cid=rs_store("cid")
rs_Store.Close

if Request.Form("update order") <> "" then
	'UPDATING ORDER DATA IN STORE_PURCHASES & STORE TRANSACTIONS TABLES
	if Request.Form("verified") <> "" then 
		verified = -1 
	else 
		verified = 0
	end if	
	
	totp = Request.Form("cnt")
	totpr = 0
	
	for i=0 to clng(totp)-1

			Quant = Request.Form("Quant_"&i)
			'rez = checkQuantity(Quant,Request.Form("IID_"&i))
			rez =""
         if rez<>"" then
				Response.redirect "error.asp?Message_Id=101&Message_Add="&rez
			else
				SKU = Request.Form("SKU_"&i)
				IName = Request.Form("IName_"&i)

				if Not IsNumeric(Request.Form("SPrice_"&i)) then
					Response.redirect "error.asp?Message_Id=101&Message_Add=Price must be numeric"
				else
					if (cdbl(Request.Form("SPrice_"&i))<0) then
						Response.redirect "error.asp?Message_Id=101&Message_Add=Price must be over 0"
					else
						if Request.Form("SPrice_"&i)<>"" then
							SPrice = Request.Form("SPrice_"&i)
						else 
							SPrice = 0
						end if
						TID = Request.Form("Transaction_Id_"&i)
						if Request.Form("Delete_"&i) = "" then
							sql_update = "update store_transactions set Quantity = "&Quant&", Item_Sku ='"&SKU&"' , Item_Name = '"&checkstringforq(IName)&"' , Sale_Price = "&SPrice&" where Transaction_Id="&TID&" and Store_Id="&Store_Id
							totpr = totpr + Quant * SPrice
						else
							sql_update = "delete from store_transactions where Transaction_Id="&TID&" and oid="&oid&";"
						end if
						conn_store.Execute sql_update
					end if
				end if
			end if
	next
	
	price_for_new = 0
	on error resume next



	if isNumeric(request.form("SKU_NEW")) then
		
			sale_price = getSalePrice(clng(request.form("SKU_NEW")))
			sql_update ="insert into store_transactions (Store_ID, OID, Shopper_ID, Quantity, Item_Id, Item_Sku, Item_Name, Sale_Price, Wholesale_Price, Item_Weight ) select "&Store_ID&", "&Oid&", (select shopper_id from store_purchases where OID="&oid&" and store_id="&store_id&"), "&request.form("Quant_NEW")&", Item_Id, Item_Sku, Item_Name, "&Sale_Price&", Wholesale_Price, Item_Weight from store_items where store_id="&store_id&" and item_id ="&request.form("SKU_NEW")& ""
			conn_store.Execute sql_update
			price_for_new = cdbl(sale_price)*cdbl(request.form("Quant_NEW"))

	elseif request.form("SKU_NEW")<>"" then
		Response.Redirect "error.asp?Message_Id=101&Message_Add="&Server.urlencode("You have not entered a valid item id.")
	end if

  sql_sel_item = "select sum(sale_price*quantity) as total from store_transactions where oid = "&Oid&" and store_id="&store_id
	rs_Store.open sql_sel_item,conn_store,1,1
  totpr = rs_store("total") + price_for_new
  rs_store.close

	gtotal = totpr
	cupa = cdbl(Request.Form("Coupon_Amount"))
	if cupa>0 then
		Response.Redirect "error.asp?Message_Id=101&Message_Add="&Server.urlencode("The coupon discount must be less then 0.  Please enter a negative number.")
	end if
	if Request.Form("Tax")<>"" then
		gtotal = gtotal + cdbl(Request.Form("Tax"))
	end if
	if Request.Form("Shipping_Method_Price")<>"" then
		gtotal = gtotal + cdbl(Request.Form("Shipping_Method_Price"))
	end if
	if Request.Form("Coupon_Amount")<>"" then
		gtotal = gtotal + cupa
	end if
	Rewards_Amount=cdbl(Request.Form("Rewards_Amount"))
	if Rewards_Amount<>0 then
		gtotal = gtotal + Rewards_Amount
	end if
     Gift_Wrapping_Amount=cdbl(Request.Form("Gift_Wrapping_Amount"))
	if Gift_Wrapping_Amount<>0 then
		gtotal = gtotal + Gift_Wrapping_Amount
	end if
	Giftcert_Amount= cdbl(Request.Form("Giftcert_Amount"))
	if Giftcert_Amount<>0 then
		gtotal = gtotal + Giftcert_Amount
	end if
	if gtotal<0 then
		gtotal = 0
	end if

	sql_update = "update Store_Purchases set " &_
		" ShipCompany = '"&checkstringforq(Request.Form("scompany"))&"' " &_
		" , ShipFirstname = '"&checkstringforq(Request.Form("sfname"))&"' " &_
		" , ShipLastname = '"&checkstringforq(Request.Form("slname"))&"' " &_
		" , ShipAddress1 = '"&checkstringforq(Request.Form("saddr1"))&"' " &_
		" , ShipAddress2 = '"&checkstringforq(Request.Form("saddr2"))&"' " &_
		" , ShipCity = '"&checkstringforq(Request.Form("scity"))&"' " &_
		" , ShipState = '"&checkstringforq(Request.Form("sstate"))&"' " &_
		" , ShipZip = '"&checkstringforq(Request.Form("szip"))&"' " &_
		" , ShipCountry = '"&checkstringforq(Request.Form("scntry"))&"' " &_
		" , ShipPhone = '"&checkstringforq(Request.Form("sphone"))&"' " &_
		" , ShipFax = '"&checkstringforq(Request.Form("sfax"))&"' " &_
		" , ShipEMail = '"&checkstringforq(Request.Form("smail"))&"' " &_
		" , verified = "&verified&"" &_
		" , Verified_Ref = '"&Request.Form("Verified_Ref")&"' " &_ 
		" , Payment_Method = '"&Replace(Request.Form("Payment_Method"),"'","''")&"' " &_ 
		" , Purchase_Date = '"&Request.Form("Purchase_Date")&"' " &_
		" , Cust_po = '"&Request.Form("cust_po")&"' " &_
		" , Shipping_Method_Name = '"&checkstringforq(Request.Form("Shipping_Method_Name"))&"' " &_
		" , CardName = '"&checkstringforq(Request.Form("CardName"))&"' " &_
		" , CardExpiration = '"&Request.Form("CardExpiration")&"' " &_
		" , AuthNumber = '"&Request.Form("AuthNumber")&"' " &_
		" , Coupon_ID = '"&checkstringforq(Request.Form("Coupon_ID"))&"' " &_
		" , Total = "&totpr&" " &_
		" , Tax = "&Request.Form("Tax")&" " &_
		" , Shipping_Method_Price = "&Request.Form("Shipping_Method_Price")&" " &_
		" , Coupon_Amount = "&(-cupa)&" " &_
		" , Giftcert_amount = "&(-Giftcert_amount)&" " &_
		" , Gift_wrapping_amount = "&(Gift_wrapping_amount)&" " &_
		" , Rewards = "&(-Rewards_Amount)&" " &_
		" , Grand_Total = "& gtotal &" "
		if strServerPort = 443 then
			sql_update = sql_update & " , CardNumber = '"&EnCrypt(Request.Form("CardNumber"))&"' "
		end if
		sql_update = sql_update & "where oid = "&Oid&" and Store_id ="&Store_id
   conn_store.Execute sql_update

	sql_update = "update Store_Customers set " &_
		" Company = '"&checkstringforq(Request.Form("bcompany"))&"' " &_
		" , First_name = '"&checkstringforq(Request.Form("bfname"))&"' " &_
		" , Last_name = '"&checkstringforq(Request.Form("blname"))&"' " &_
		" , Address1 = '"&checkstringforq(Request.Form("baddr1"))&"' " &_
		" , Address2 = '"&checkstringforq(Request.Form("baddr2"))&"' " &_
		" , City = '"&checkstringforq(Request.Form("bcity"))&"' " &_
		" , State = '"&checkstringforq(Request.Form("bstate"))&"' " &_
		" , Zip = '"&checkstringforq(Request.Form("bzip"))&"' " &_
		" , Country = '"&checkstringforq(Request.Form("bcntry"))&"' " &_
		" , Phone = '"&checkstringforq(Request.Form("bphone"))&"' " &_
		" , Fax = '"&checkstringforq(Request.Form("bfax"))&"' " &_
		" , EMail = '"&checkstringforq(Request.Form("bmail"))&"' " &_
		
		"where Cid="&cid&" AND Record_type = 0"
	conn_store.Execute sql_update


end if

'LOAD VALUES FROM STORE_PURCHASES TABLE
sql_select_purchases = "Select * from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id 
rs_Store.open sql_select_purchases,conn_store,1,1
	Verified = Rs_store("Verified")
	if Verified = -1 then
		checked = "checked"
	end if
	Cid=rs_store("cid")
	CCid = rs_store("ccid")
	Verified_Ref = Rs_store("Verified_Ref")
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Shipping_Method_Name = Rs_store("Shipping_Method_Name")
	Tax = Rs_store("Tax")
	Total = Rs_store("Total")
	Coupon_id = Rs_store("Coupon_id")
	Coupon_Amount = Rs_store("Coupon_Amount")
	giftcert_id = Rs_store("giftcert_id")
     giftcert_Amount = Rs_store("giftcert_amount")
     rewards_Amount = Rs_store("rewards")
     Gift_Wrapping = Rs_store("gift_wrapping")
     Gift_Wrapping_Amount = Rs_store("gift_wrapping_amount")
	Grand_Total = Rs_store("Grand_Total")
	Payment_Method = Rs_store("Payment_Method")
	CardName = Rs_store("CardName")
	If strServerPort<> 443 then
	  BeginCardNumber = Left(DeCrypt(Rs_store("CardNumber")),4)
	  EndCardNumber = Right(DeCrypt(Rs_store("CardNumber")),4)
	  CardNumber = BeginCardNumber & "**********" & EndCardNumber
	else
	  CardNumber = DeCrypt(Rs_store("CardNumber"))
	End if
	CardExpiration = Rs_store("CardExpiration")
	AuthNumber = Rs_store("AuthNumber")
	Purchase_Date = Rs_store("Purchase_Date")
	Cust_Po = Rs_store("Cust_PO")
	Cust_Notes = Rs_store("Cust_Notes")
	Is_Gift = Rs_store("Is_Gift")
	Gift_Wrapping = Rs_store("Gift_Wrapping")
	Gift_Message = Rs_store("Gift_Message")
	Coupon_Amount = Rs_store("Coupon_Amount")
	giftcert_id = Rs_store("giftcert_id")
     giftcert_Amount = Rs_store("giftcert_amount")
     rewards_Amount = Rs_store("rewards")
     Gift_Wrapping = Rs_store("gift_wrapping")
     Gift_Wrapping_Amount = Rs_store("gift_wrapping_amount")
	
	if Is_Gift = -1 then
		Is_Gift_Text = "Yes"
	else
		Is_Gift_Text = "No"
	end if
	if Gift_Wrapping = -1 then
		Gift_Wrapping = "Yes"
	else
		Gift_Wrapping = "No"
	end if
	ShipCompany =Rs_store("ShipCompany")
	ShipFirst_name=Rs_store("ShipFirstname")
	ShipLast_name =Rs_store("ShipLastname")
	ShipAddress1 =Rs_store("ShipAddress1")
	ShipAddress2 =Rs_store("ShipAddress2")
	ShipCity=Rs_store("ShipCity")
	ShipState=Rs_store("ShipState")
	Shipzip =Rs_store("Shipzip")
	ShipCountry= Rs_store("ShipCountry")
	ShipPhone  =Rs_store("ShipPhone")
	ShipFax =Rs_store("ShipFax")
	ShipEMail=Rs_store("ShipEMail")
	if Rs_store("Verified") = -1 then
		Verified = "Verified" 
	Else 
		Verified = "Not Verified" 
	end if
rs_Store.Close



addPicker = 1
sFormAction = "order_edit.asp?oid="&oid&"&cid="&cid
sName = "ORD_EDIT"
sTitle = "Edit Order - #"&Oid
sFullTitle = "<a href=orders.asp class=white>Orders</a> > Edit - #"&Oid
sSubmitName = "Store_Activation_Update"
thisRedirect = "order_edit.asp"
sMenu = "orders"
createHead thisRedirect
%>
 


 


				<tr bgcolor='#FFFFFF'><td colspan=3><a href="order_details.asp?Id=<%= oid %>" class=link><b>View Order Details</b></a></td></tr>
				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Order Verified</td>
					<td class="inputvalue">
						<input class="image" type="checkbox" <%= checked %> name="Verified" value="1">
										<% small_help "Order Verified" %></td>
				</tr>
				
				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Verification Code</td>
					<td class="inputvalue">
						<input type="text" name="Verified_Ref" value="<%= Verified_Ref %>" size="20">
										<% small_help "Verification Code" %></td>
				</tr>
  
				<tr bgcolor='#FFFFFF'>
					<td colspan=3 align=center>
						<input type="submit" class="Buttons" value="Update Order" name="Update Order"><BR>
                                                </td>
				</tr>
				
				<tr bgcolor='#FFFFFF'>
					<td colspan=3><hr width="600" align="left"></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td height="508" colspan=3>
						<table width=600 border=0 cellpadding=2 cellspacing=0>
							
							<tr bgcolor='#FFFFFF'>
								<td><B>Customer ID</B></td>
								<td><%= CCid %></td>
								<td><B>Date</B></td>
								<td>
									<input type="text" size="20" name="Purchase_Date" value="<%= FormatDateTime(Purchase_Date) %>"></td>
								<td><b>Purchase Order</b></td>
								<td>
									<input type="text" size="10" name="Cust_PO" value="<%= Cust_PO %>"></td>
							</tr>
							
							<tr bgcolor='#FFFFFF'> 
								<td><B>Order ID</B></td>
								<td><%= Oid %></td>
								<td><B>Terms</B></td>
								<td>
									<input type=text name="Payment_Method" value='<%=Payment_Method%>'>

									</td>
								<td colspan="2">&nbsp;</td>
							</tr>
						</table>
						
						<% if Payment_Method = "Visa" or Payment_Method = "Mastercard" or Payment_Method = "Discover" or Payment_Method = "American Express" or Payment_Method = "Diner's Club" then %>
							<hr width="600" align="left">
							<table width=600 border=0 cellpadding=2 cellspacing=0>

								<tr bgcolor='#FFFFFF'>
									<td colspan="3"><B>This is a credit card order</B></td>
								</tr>

								<tr bgcolor='#FFFFFF'>
									<td>&nbsp;</td>
									<td><B>Authorization Number</B></td>
									<td>
										<input type="text" size="20" name="AuthNumber" value="<%= AuthNumber %>"></td>
								</tr>
								
								<tr bgcolor='#FFFFFF'>
									<td>&nbsp;</td>
									<td><B>Card Name</B></td>
									<td>
										<input type="text" size="20" name="CardName" value="<%= CardName %>"></td>
								</tr>
	 
								<tr bgcolor='#FFFFFF'>
									<td>&nbsp;</td>
									<td><B>Card Number</B></td>
									<td>
										<input type="text" size="20" name="CardNumber" value="<%= CardNumber %>">
										<% if Port <> 443 then %>
										<font size=1>To protect your customers, credit card numbers can't be seen or edited unless you are logged in securely.</font>
										<% end if %>
										</td>
								</tr>
							
								<tr bgcolor='#FFFFFF'>
									<td>&nbsp;</td>
									<td><b>Card Expiration</b></td>
									<td>
										<input type="text" size="20" name="CardExpiration" value="<%= CardExpiration %>"></td>
								</tr>
							</table>
						<% End If %>	
						
						<hr width="600" align="left">
						
						<table width="600" border="0" cellpadding="2" cellspacing="0">

							<tr bgcolor='#FFFFFF'>
								<td valign="top" width="50%"><b>Bill To</b>
									<% sql_select_cust =  "SELECT Store_Customers.Last_name, Store_Customers.First_name, Store_Customers.Company, Store_Customers.Address1, Store_Customers.Address2, Store_Customers.City, Store_Customers.State, Store_Customers.Zip, Store_Customers.Country, Store_Customers.EMail, Store_Customers.Phone,Store_Customers.FAX FROM Store_Customers WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type = 0"  %>
									<% rs_Store.open sql_select_cust,conn_store,1,1 %>
									<% rs_Store.MoveFirst %>

										<table cellpadding="0" cellspacing="5" border="0">
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Company</td>
												<td class="inputvalue">
													<input type="text" name="bcompany" size="25" value="<%= rs_Store("Company") %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">First Name</td>
												<td class="inputvalue">
													<input type="text" name="bfname" size="25" value="<%= rs_Store("First_name")	%>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Last Name</td>
												<td class="inputvalue">
													<input type="text" name="blname" size="25" value="<%= rs_Store("Last_name") %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Address Line 1</td>
												<td class="inputvalue">
													<input type="text" name="baddr1" size="25" value="<%= rs_Store("Address1") %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Address Line 2</td>
												<td class="inputvalue">
													<input type="text" name="baddr2" size="25" value="<%= rs_Store("Address2") %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">City</td>
												<td class="inputvalue">
													<input type="text" name="bcity" size="25" value="<%=	rs_Store("City") %>"></td>
											</tr>
											

				                                                        <tr bgcolor='#FFFFFF'>
												<td class="inputname">State</td>
												<td class="inputvalue">
													<input type="text" name="bstate" size="25" value="<%= rs_Store("state") %>"></td>
											</tr>
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Zip</td>
												<td class="inputvalue">
													<input type="text" name="bzip" size="25" value="<%= rs_Store("zip") %>"></td>
											</tr>
		
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Country</td>
												<td class="inputvalue"><%= create_country_list ("bcntry",rs_Store("Country"),1,"") %></td>
											</tr>

											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Phone</td>
												<td class="inputvalue">
													<input type="text" name="bphone" size="25" value="<%= rs_Store("Phone") %>"></td>
											</tr>
		  
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Fax</td>
												<td class="inputvalue">
													<input type="text" name="bfax" size="25" value="<%= rs_Store("Fax") %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Email</td>
												<td class="inputvalue">
													<input type="text" name="bmail" size="25" value="<%= rs_Store("EMail") %>"></td>
											</tr>
										</table>

									<% rs_Store.Close %>
								</td>
								
								<td width="50%" valign="top"><b>Ship To</b>

										<table cellpadding="0" cellspacing="5" border="0">

											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Company</td>
												<td class="inputvalue">
													<input type="text" name="scompany" size="25" value="<%= ShipCompany %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">First Name</td>
												<td class="inputvalue">
													<input type="text" name="sfname" size="25" value="<%= ShipFirst_name %>"></td>
											</tr>
		  
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Last Name</td>
												<td class="inputvalue">
													<input type="text" name="slname" size="25" value="<%= ShipLast_name %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Address Line 1</td>
												<td class="inputvalue">
													<input type="text" name="saddr1" size="25" value="<%= ShipAddress1 %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Address Line 2</td>
												<td class="inputvalue">
													<input type="text" name="saddr2" size="25" value="<%= ShipAddress2 %>"></td>
											</tr>
												
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">City</td>
												<td class="inputvalue">
													<input type="text" name="scity" size="25" value="<%= ShipCity %>"></td>
											</tr>
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">State</td>
												<td class="inputvalue">
													<input type="text" name="sstate" size="25" value="<%= Shipstate %>"></td>
											</tr>


											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Zip</td>
												<td class="inputvalue">
													<input type="text" name="szip" size="25" value="<%= Shipzip %>"></td>
											</tr>
		
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Country</td>
												<td class="inputvalue"><%= create_country_list ("scntry",ShipCountry,1,"") %></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Phone</td>
												<td class="inputvalue">
													<input type="text" name="sphone" size="25" value="<%= ShipPhone %>"></td>
											</tr>
		  
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Fax</td>
												<td class="inputvalue">
													<input type="text" name="sfax" size="25" value="<%= ShipFax %>"></td>
											</tr>
											
											<tr bgcolor='#FFFFFF'>
												<td class="inputname">Email</td>
												<td class="inputvalue">
													<input type="text" name="smail" size="25" value="<%= ShipEMail %>"></td>
											</tr>
										</table>
									
								</td>
							</tr>
						</table>
						
						<hr width="600" align="left">
	 
						
						Order Details
	  
						<table width="600" border="1" cellpadding="0" cellspacing="0">
	  
						<tr bgcolor='#FFFFFF'>
							<td align="center" width="10%"><b>Quantity</b></td>
							<td align="center" width="30%"><b>Item SKU</b></td>
							<td align="center" width="30%"><b>Item Name</b></td>
							<td align="center" width="10%"><b>Sale Price</b></td>
							<td align="center" width="10%"><b>Total Price</b></td>
							<td align="center" width="10%"><b>Delete</b></td>
						</tr>
	  
						<% sql_select = "select * from store_transactions where oid = "&Oid&" and Store_id ="&Store_id %>
						<% rs_Store.open sql_select,conn_store,1,1 %>
						<% cnt = 0 %>
						<% do while not rs_store.eof %>
							<tr bgcolor='#FFFFFF'>
								<input type="Hidden" name="IID_<%= cnt %>" value="<%= rs_store("Item_ID") %>">
								<td><input size="5" type="text" name="Quant_<%= cnt %>" value="<%= rs_store("Quantity") %>"></td>
								<td><%= rs_store("Item_sku") %>
								<input type="hidden" name="SKU_<%= cnt %>" value="<%= rs_store("Item_sku") %>"></td>
								<td><input size="20" type="text" name="IName_<%= cnt %>" value="<%= rs_store("Item_name") %>"></td>
								<td><input size="10" type="text" name="SPrice_<%= cnt %>" value="<%= rs_store("Sale_price") %>"></td>
								<td align="right"><%= Currency_Format_Function(rs_store("Sale_price")*rs_store("Quantity")) %></td>
								<input type="hidden" name="Transaction_Id_<%= cnt %>" value="<%= rs_store("Transaction_Id") %>">
								<td align="center"><input class="image" type="checkbox" name="Delete_<%= cnt %>"></td>
							</tr>
							<% cnt = cnt + 1 %>
							<% rs_store.movenext %>
						<% loop %>
						<% rs_store.close %>
						<input type="hidden" name="cnt" value="<%= cnt %>">
						
						<tr bgcolor='#FFFFFF'>
							<td colspan="6"><hr></td>
						</tr>

						<tr bgcolor='#FFFFFF'>
							<td colspan="6">Add Item to Order</td>
						</tr>
	 
						<tr bgcolor='#FFFFFF'>

							<td><input size="5" type="text" name="Quant_NEW" value=1></td>
							<td><input type="text" name="SKU_NEW" value="0" size=15></td>
							<td colspan="3"><a class="link" href="JavaScript:goItemPicker('SKU_NEW');">Select Item</a><br>
								</td>
						</tr>
					</table>
					
					<br>
		
					<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
						
						<tr bgcolor='#FFFFFF'> 
							<TD width="301" height="1" colspan="4"><STRONG>Total</STRONG></TD>
						<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Total) %></STRONG></TD>
						</tr>
	
						<tr bgcolor='#FFFFFF'>
							<TD width="301" height="1" colspan="4"><STRONG>Shipping Method</STRONG></TD>
							<TD width="71" height="1"><STRONG>
								<input type="text" name="Shipping_Method_Name" size="10" value="<%= Shipping_Method_Name %>"></STRONG></TD>
					</tr>
	
						<tr bgcolor='#FFFFFF'>
							<TD width="301" height="1" colspan="4"><STRONG>Shipping Price</STRONG>	</TD>
						<TD width="71" height="1"><STRONG>
								<input type="text" name="Shipping_Method_Price" size="10" value="<%= Shipping_Method_Price %>"></STRONG></TD>
						</tr>
						
						<tr bgcolor='#FFFFFF'> 
							<TD width="301" height="1" colspan="4"><STRONG>Tax</STRONG></TD>
							<TD width="71" height="1"><STRONG>
								<input type="text" name="tax" size="10" value="<%= Tax %>"></STRONG></TD>
						</tr>


							<tr bgcolor='#FFFFFF'>
								<td align="left" colspan="4"><STRONG><i>Coupon ID</i></STRONG> </td>
								<td align="left"><STRONG>
									<input type="text" size="10" name="Coupon_Id" value="<%= Coupon_Id %>"></td>
							</tr>

							<tr bgcolor='#FFFFFF'>
								<td align="left" colspan="4"><STRONG><i>Coupon</i></STRONG> </td>
								<td align="left"><STRONG>
									<input type="text" size="10" name="Coupon_Amount" value="<%= -Coupon_Amount %>"></td>
							</tr>
							<tr bgcolor='#FFFFFF'>
								<td align="left" colspan="4"><STRONG><i>Rewards</i></STRONG> </td>
								<td align="left"><STRONG>
									<input type="text" size="10" name="rewards_Amount" value="<%= -rewards_Amount %>"></td>
							</tr>
							<tr bgcolor='#FFFFFF'>
								<td align="left" colspan="4"><STRONG>Gift Wrapping</STRONG> </td>
								<td align="left"><STRONG>
									<input type="text" size="10" name="Gift_Wrapping_amount" value="<%= Gift_Wrapping_amount %>"></td>
							</tr>
							<tr bgcolor='#FFFFFF'>
								<td align="left" colspan="4"><STRONG>Gift Certificate</STRONG> </td>
								<td align="left"><STRONG>
									<input type="text" size="10" name="Giftcert_Amount" value="<%= -Giftcert_Amount %>"></td>
							</tr>
						
						
						<tr bgcolor='#FFFFFF'>
							<TD width="304" height="1" colspan="4"><STRONG>Grand Total</STRONG> </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Grand_Total) %></STRONG></TD>
						</tr>
					</TABLE>
				</td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
				<td height="19" colspan=5 class=instructions>
					<input type="submit" class="Buttons" value="Update Order" name="Update Order"><BR><Font class=small>Please hit update a second time after adding/removing your final item to recalculate the correct final total.</font></td>
			</tr>
		<% createFoot thisRedirect, 0%>
