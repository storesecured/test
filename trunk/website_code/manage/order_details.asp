<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="pagedesign.asp"-->
<%

'LOAD ORDER DETAILS FOR SPECIFIED ORDER FROM THE DATABASE
Oid = Request.QueryString("Id")
if Oid="" then
   Oid = Request.QueryString("Oid")
end if
if Oid="" or not isNumeric(oid) then
   response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("You must choose a valid order to view.<BR><BR><a href=orders.asp>Click here to return to the order list.</a>")
end if

'CHECK IF MERCHANT WANT'S TO UPDATE UPDATE_VERIFICATION_STATUS
if Request.Form("Update_Verification_status") <> "" then
	Verified_Ref = Replace(Request.Form("Verified_Ref"), "'", "''")
	if Request.Form("verified") <> "" then
		verified = -1 
	else
		verified = 0
	end if	
	sql_update = "update Store_Purchases set verified = "&verified&", Verified_Ref = '"&Verified_Ref&"' where oid = "&Oid&" and Store_id ="&Store_id
	conn_store.Execute sql_update
	if Request.Form("update_stock") = "-1" then
		'UPDATE THE QUANTITIES, USED TO TRACK INVENTORY
		  sql_sel_loop = "SELECT Item_Id, SUM(Quantity) AS Quantity_total FROM Store_Transactions WITH (NOLOCK) GROUP BY Item_Id, OID, Store_ID HAVING (OID = "&OID&") AND (Store_ID = "&Store_ID&")"
		  set myfields=server.createobject("scripting.dictionary")
		  Call DataGetrows(conn_store,sql_sel_loop,mydata,myfields,noRecords)

		  if noRecords = 0 then
		  FOR rowcounter= 0 TO myfields("rowcount")
			  sql_update="UPDATE Store_Items SET Store_Items.Quantity_in_stock = [Store_Items].[Quantity_in_stock]- "&mydata(myfields("quantity_total"),rowcounter)&" where Store_Items.Item_Id = "&mydata(myfields("item_id"),rowcounter)&" and Store_Items.Store_id ="&Store_id
			  conn_store.Execute sql_update
		  Next
		  End If
		  'set a flag in the db to show that stock has been updated for this order.
		  sql_update = "update Store_Purchases set Stock_Updated = -1 where oid = "&Oid&" and Store_id ="&Store_id
		  conn_store.Execute sql_update

	end if
end if

if Request.Form("Capture") <> "" or Request.Form("Credit") <> "" or Request.Form("Void") <> "" then
	  %>
          <!--#include file="include/gateways_functions.asp"-->
          <%
          if Request.Form("Capture") <> "" then
	sType = "Capture"
	trans_type = "2"
	  elseif Request.Form("Credit") <>"" then
	sType = "Credit"
	trans_type = "3"
	elseif Request.Form("Void") <> "" then
	       sType ="Void"
	trans_type = "3"
	  end if
	  Real_Time_Processor = Request.form("Processor_id")
	  GGrand_Total = Request.Form("new_amount")
	  if Real_Time_Processor = 1 then
		%><!--#include file="include\plugnpay\pnp.asp"--><%
	elseif Real_Time_Processor = 2 then
		%><!--#include file="include\Authorizenet\Authorizenet.asp"--><%
	elseif Real_Time_Processor = 7 then
	  %><!--#include file="include\Echo\Echo.asp"--><%
	  elseif Real_Time_Processor = 10 then
	  %><!--#include file="include\linkpoint\linkpoint.asp"--><%
	  elseif Real_Time_Processor = 36 then
	  %><!--#include file="include\paypalpro\paypalpro.asp"--><%
	  elseif Real_Time_Processor = 38 then
           %>

           <!--#include file="include\google\google.asp"-->
          <%
	  else
	       response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("Your processor is not yet supported for automatic capture.")
          end if
	sql_update_purchases = "update store_purchases set Transaction_Type='"&trans_type&"' where (oid = "&oid&" or masteroid="&oid&") AND Store_id ="&Store_id
	conn_store.Execute sql_update_purchases
end if

'LOAD VALUES FROM STORE_PURCHASES TABLE
sql_select_purchases = "Select * from Store_Purchases WITH (NOLOCK) where (oid = "&Oid&" or masteroid="&oid&") and Store_id ="&Store_id

GGrand_total=0
rs_Store.open sql_select_purchases,conn_store,1,1
do while not rs_store.eof
	rsoid = Rs_store("oid")
	rsGrand_Total = Rs_store("Grand_Total")
	if formatnumber(rsoid)=formatnumber(oid) then
        Grand_Total = Rs_store("Grand_Total")
        Verified = Rs_store("Verified")
	if Verified = -1 then
		checked = "checked"
	end if
	shopper_id = rs_store("shopper_id")
	maxmind_score = rs_store("maxmind_score")
        cid = rs_store("cid")
	ccid = rs_store("ccid")
	Verified_Ref = Rs_store("Verified_Ref")
	Masteroid = Rs_store("Masteroid")
	Invoice_Id = Rs_Store("Invoice_Id")
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Shipping_Method_Name = Rs_store("Shipping_Method_Name")
	Coupon_id = Rs_store("coupon_id")
     Coupon_Amount = Rs_store("coupon_amount")
     giftcert_id = Rs_store("giftcert_id")
     giftcert_Amount = Rs_store("giftcert_amount")
     rewards_Amount = Rs_store("rewards")
     Gift_Wrapping = Rs_store("gift_wrapping")
     Gift_Wrapping_Amount = Rs_store("gift_wrapping_amount")
	Tax = Rs_store("Tax")
	Total = Rs_store("Total")
	Payment_Method = Rs_store("Payment_Method")
	Purchase_Completed = Rs_store("Purchase_Completed")
	Last_Step = Rs_store("Last_Step")
	CardName = Rs_store("CardName")
	Google_FinancialStatus = Rs_store("Google_FinancialStatus")
        Google_FulfillmentStatus = Rs_store("Google_FulfillmentStatus")
	If strServerPort<> 443 then
	  if Rs_store("CardNumber")<>"" then
       BeginCardNumber = Left(trim(DeCrypt(Rs_store("CardNumber"))),4)
  	    EndCardNumber = Right(trim(DeCrypt(Rs_store("CardNumber"))),4)
  	    CardNumber = BeginCardNumber & "******" & EndCardNumber
  	    CardNumber = CardNumber&"<font size=1><BR><a href=https://"&request.servervariables("HTTP_HOST")&"/order_details.asp?Id="&request.querystring("Id")&">Switch to Secure mode to see the entire number</a></font>"
  	  else
  	    CardNumber = "Unknown"
     end if
	elseif Session("Super_User") = 1 then
		CardNumber = "Hidden"
	else
	  CardNumber = DeCrypt(Rs_store("CardNumber"))
	End if
	CardExpiration = Rs_store("CardExpiration")
	IssueNumber = Rs_store("IssueNumber")
	IssueStart = Rs_store("IssueStart")
	AuthNumber = Rs_store("AuthNumber")
	Verified_Ref = Rs_store("Verified_Ref")
	AVS_Result = Rs_store("AVS_Result")
	Card_Code_Verification = Rs_store("Card_Code_Verification")
	Processor_id = Rs_store("Processor_id")
	Transaction_Type = Rs_store("Transaction_Type")
	Purchase_Date = Rs_store("Purchase_Date")
	Cust_Po = Rs_store("Cust_PO")
	Cust_Notes = Rs_store("Cust_Notes")
	Referer = Rs_store("Referer")
	ClientIp = Rs_store("ClientIp")
	Custom_fields = Rs_store("Custom_fields")
	Is_Gift = Rs_store("Is_Gift")
	Gift_Message = Rs_store("Gift_Message")
	if Is_Gift = -1 then
		Is_Gift_Text = "Yes"
	else
		Is_Gift_Text = "No"
	end if
	if Gift_Wrapping = -1 then
		Gift_Wrapping1 = "Yes"
	else
		Gift_Wrapping1 = "No"
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
	ShipResidential=Rs_Store("ShipResidential")
	if Rs_store("Verified") = -1 then
		Verified = "Verified" 
	Else 
		Verified = "Not Verified" 
	end if
	Stock_Updated = Rs_store("Stock_Updated")
	Recurring_Total = rs_store("Recurring_Total")
	Recurring_FeeT = rs_store("Recurring_Fee")
	Recurring_Tax = rs_store("Recurring_Tax")
	Recurring_Days = rs_store("Recurring_Days")
	end if
	rs_store.movenext
	GGrand_total=GGrand_total+rsGrand_Total
loop
rs_Store.Close

sFormAction = "order_details.asp?Id="&oid
sName = "Order #"&Oid&"&nbsp;Details"
sFormName = "activation"
sTitle = "Order Details - #"&oid
sFullTitle = "<a href=orders.asp class=white>Order</a> > Details - #"&oid
sSubmitName = "Store_Activation_Update"
thisRedirect = "order_details.asp"
sMenu = "orders"
createHead thisRedirect

on error resume next
 %>

<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
function confirmPayment()
{
var agree=confirm("Are you sure?");
if (agree)
	return true ;
else
	return false ;
}
function fnAddNotes(orderID,storeID)
{
	window.open("addOrderNotes.asp?oID="+orderID,"AddNotes","top=10,left=50,width=600,height=400,status=yes,resize=no,scrollbars=1");
}
</script>


	 <% if purchase_completed=0 then %>
	 <TR bgcolor=red><td colspan=3 align=center><font color=white size=3><B>Abandoned</b></font></td></tr>
	 <tr bgcolor=FFFFFF><td colspan=3>
	 <% if Last_Step="payment.asp" then %>
	 <a class=link target=_blank href="https://<%= Secure_Name %>/include/process_manual.asp?Shopper_Id=<%=shopper_id%>&Return_From=Admin&Oid=<%= Oid %>">Manually Complete Order</a> (This will not charge the customer)
	 <% end if %>
	 <a class=link target=_blank href="<%= Site_Name %>show_big_cart.asp?Shopper_Id=<%=shopper_id%>">Open Customers Shopping Cart</a>
	 </td></tr>

	<% elseif verified="Verified" then %>
	<TR bgcolor=green><td colspan=3 align=center><font color=white size=3><B>Paid</b></font></td></tr>
	<% else %>
	<TR bgcolor=yellow><td colspan=3 align=center><font color=black size=3><B>Unpaid</b></font></td></tr>
	<% end if %>
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="74">
			<table width="600" height="593">
				<% if masteroid<>"" then %>
            <tr bgcolor='#FFFFFF'><TD bgcolor=red width=100% colspan=3><B><font color=white>This order is related to Master Order Id <a class=link href=order_details.asp?Id=<%= masteroid %>><%= masteroid %></a></font></b></TD></TR>
				<% end if %>
            <tr bgcolor='#FFFFFF'>
				<td colspan=2><a href="order_printable.asp?Oid=<%= oid%>&cid=<%= cid %>" class=link target=_blank><b>Printable Copy Details</b></a> &nbsp; &nbsp;
					<a href="order_printable.asp?Oid=<%= oid%>&cid=<%= cid %>&Cart_Type=Invoice" class=link target=_blank><b>Printable Copy Invoice</b></a> &nbsp; &nbsp;
                                        <a href="order_packing.asp?Oid=<%= oid%>&cid=<%= cid %>&Cart_Type=Packing" class=link target=_blank><B>Packing Slip</b></a> &nbsp; &nbsp;
                                     <a href=https://<%= Secure_Name %>/include/notification.asp?Oid=<%= oid %>&Shopper_Id=<%= shopper_id %>&Cid=<%=cid %> class=link target=_blank><B>Resend Notification</b></a> &nbsp; &nbsp;
                                     <a href="order_edit.asp?Id=<%= oid %>" class=link><b>Edit Order</b></a>
				</td></tr>
				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Verified</td>
					<td class="inputvalue">
						<input class="image" type="checkbox" <%= checked %> name="Verified" value="-1">
					<% small_help "Verified" %></td>
				</tr>
 
				<tr bgcolor='#FFFFFF'>
				<td class="inputname">Verification Code</td>
				<td class="inputvalue"><input type="text" name="Verified_Ref" value="<%= Verified_Ref %>" size="60">
			 <% small_help "Verification Code" %></td>
				</tr>
				<% if Stock_Updated = -1 then %>
				<tr bgcolor='#FFFFFF'>
					<td height="19">Inventory Stock Already Updated</td>
				</tr>
				<% else %>
				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Update Inventory Stock</td>
					<td class="inputvalue">
						<input class="image" type="checkbox" name="update_stock" value="-1"></td>
				</tr>
				<% end if %>
				<tr bgcolor='#FFFFFF'>
					<td colspan=2 align=center>
						<input type="submit" class="Buttons" value="Update Order" name="Update_Verification_status"></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td height="19" colspan=2>
					<hr width="600" align="left">
						<b>

						<a class="link" href="ship_orders.asp?oid=<%= oid %>&cid=<%= cid %>">Ship Order</a></b>&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;
						
						<b><a class="link" href="return_orders.asp?oid=<%= oid %>&cid=<%= cid %>">Return Order</a></b>
				      &nbsp;&nbsp;&nbsp;&nbsp;
					  <b><a href="Javascript:fnAddNotes('<%=Oid%>')" class=link>Add Notes</a></b>
					  &nbsp;&nbsp;&nbsp;&nbsp;
					  <b><a class="link" href="orders.asp?delete_id=<%= oid %>">Delete Order</a></b></td>

            </tr>
				
				<tr bgcolor='#FFFFFF'>
				<td height="508" colspan=2>
						<table width=600 border=0 cellpadding=2 cellspacing=0>
							
							<tr bgcolor='#FFFFFF'>
							<td><B>Date:</B> <%= FormatDateTime(Purchase_Date,2) %>&nbsp;<%= FormatDateTime(Purchase_Date,4) %> PST</td>
							<td><B>Status:</B> <%= Verified %></td>
							<td><B>Customer ID:</B> <%= CCid %></td>
							</tr>
							
							<tr bgcolor='#FFFFFF'>
							<td><B>Order ID:</B> <%= Oid %></td>
								<td><B>Terms:</B> <%= Payment_Method %></td>
								<td><b>Purchase Order:</b> <%= Cust_PO %></td>
							</tr>
							<% if not isNull(Invoice_Id) then %>
							<tr bgcolor='#FFFFFF'>
							<td><B>Invoice ID:</B> <%= Invoice_ID %></td>

							</tr>

							<% end if %>

						
						</table>
						
						<% sPayments="Google Checkout,Visa,Mastercard,American Express,Discover,Diners Club,eCheck,Debit Card,PayPal-ExpressCheckout,Checks by Net,Solo,Switch,Maestro"

						if Is_In_Collection(sPayments,Payment_Method,",") and isNull(masteroid) then
							if Processor_id = 1 then
								Processor = "PlugnPay"
							elseif Processor_id = 2 then
								Processor = "Authorize.Net"
							elseif Processor_id = 4 then
								Processor = "Paypal.com"
							elseif Processor_id = 5 then
								Processor = "PsiGate"
							elseif Processor_id = 6 then
								Processor = "iTransact"
							elseif Processor_id = 7 then
								Processor = "Echo"
							elseif Processor_id = 8 then
								Processor = "Payflow Link"
							elseif Processor_id = 9 then
								Processor = "Payflow Pro"
							elseif Processor_id = 10 then
								Processor = "Linkpoint"
							elseif Processor_id = 11 then
								Processor = "2Checkout"
							elseif Processor_id = 12 then
								Processor = "Internet Secure"
							elseif Processor_id = 13 then
								Processor = "PaySystems"
							elseif Processor_id = 14 then
								Processor = "BluePay"
							elseif Processor_id = 15 then
								Processor = "Electronic Transfer"
							elseif Processor_id = 16 then
								Processor = "PayReady"
							elseif Processor_id = 17 then
								Processor = "WorldPay"
							elseif Processor_id = 18 then
								Processor = "Bank of America"
							elseif Processor_id = 19 then
								Processor = "Protx"
							elseif Processor_id = 20 then
								Processor = "Moneybookers"
							elseif Processor_id = 21 then
								Processor = "eWay"
							elseif Processor_id = 22 then
								Processor = "NoChex"
							elseif Processor_id = 23 then
								Processor = "SecPay"
							elseif Processor_id = 24 then
								Processor = "PRI"
							elseif Processor_id = 26 then
								Processor = "EFT"
							elseif Processor_id = 27 then
								Processor = "iKobo"
							elseif Processor_id = 28 then
								Processor = "Cybersource"
							elseif Processor_id = 29 then
								Processor = "Xor"
							elseif Processor_id = 30 then
								Processor = "ViaKlix"
							elseif Processor_id = 31 then
								Processor = "Propay"
							elseif Processor_id = 32 then
								Processor = "Transecute"
							elseif Processor_id = 33 then
								Processor = "ChecksByNet"
							elseif Processor_id = 35 then
								Processor = "Swissnet"
							elseif Processor_id = 36 then
								Processor = "Paypal Pro"
							elseif Processor_id = 37 then
								Processor = "Paystation"
							elseif Processor_id = 38 then
								Processor = "Google Checkout"
							elseif Processor_id = 0 or Processor_Id = "" or isNull(Processor_id) then
								Processor = "None"
							else
								Processor = Processor_id
					end if

							if AVS_Result = "A" then
								AVS = "Address matches, Zip does not"
							elseif AVS_Result = "B" then
								AVS = "Address information not provided"
							elseif AVS_Result = "E" then
								AVS = "AVS error"
							elseif AVS_Result = "G" then
								AVS = "Non-US Card Issuing Bank"
							elseif AVS_Result = "N" then
								AVS = "No match on address or zip."
							elseif AVS_Result = "P" then
								AVS = "AVS not applicable for this transaction"
							elseif AVS_Result = "R" then
								AVS = "Retry - System unavailable or timed out."
							elseif AVS_Result = "S" then
								AVS = "Service not supported by issuer"
							elseif AVS_Result = "U" then
								AVS = "Address information is unavailable"
							elseif AVS_Result = "W" then
								AVS = "9 digit zip matches, address does not."
							elseif AVS_Result = "X" then
								AVS = "Address and 9 digit zip match"
							elseif AVS_Result = "Y" then
								AVS = "Address and 5 digit zip match"
							elseif AVS_Result = "Z" then
								AVS = "5 digit zip matches, address does not"
							end if

							if Transaction_Type = "1" then
								Trans_Type = "Auth and Deposit"
							elseif Transaction_Type = "0" then
								Trans_Type = "Auth Only"
							elseif Transaction_Type = "2" then
								Trans_Type = "Delayed Capture"
							elseif Transaction_Type = "3" then
								Trans_Type = "Credited"
							end if

                     Card_Code_Verification=replace(Card_Code_Verification," ","")
							if Card_Code_Verification = "M" then
								Card_Code = "Card code matched"
							elseif Card_Code_Verification = "N" then
								Card_Code = "Card code does not match"
							elseif Card_Code_Verification = "P" then
								Card_Code = "Card code was not processed"
							elseif Card_Code_Verification = "S" then
								Card_Code = "Card code should be on card but was not indicated"
							elseif Card_Code_Verification = "U" then
								Card_Code = "Issuer was not certified for card code"
							end if

							%>

						<hr width="600" align="left">
							<table width=600 border=0 cellpadding=5 cellspacing=0>
								<tr bgcolor='#FFFFFF'>
									<td rowspan=2 width=15%><B>Processed by <%= Processor %></b></td>
									<td width=25%><B>Card Number</B><BR><%= CardNumber %></td>
									<td width=25%><B>Card Name</B><BR><%= CardName %></td>
									<td width=10%><B>Auth #</B><BR><%= AuthNumber %></td>
									<td width=25%><b>Card Exp</b><BR><%= CardExpiration %></td>

								</tr>
								<tr bgcolor='#FFFFFF'>

									<td><B>AVS Result</B><BR><%= AVS %></td>
									<td><B>Code Result</B><BR><%= Card_Code %></td>
									<td><B>Ref #</B><BR><%= Verified_Ref %></td>
									<td><b>Type</b><BR><%= Trans_Type %></td>
						 <input type=hidden name="Processor_id" value="<%= Processor_id %>">
								</tr>
									<% if Processor_id=38 then %>
								<tr bgcolor='#FFFFFF'>

									<td></td>
									<td><B>Financial Status</B><BR><%= Google_FinancialStatus %></td>
									<td><B>Fulfillment Status</B><BR><%= Google_FulfillmentStatus %></td>
								</tr>
									<% end if %>
								<% if Payment_Method="Solo" or Payment_Method="Switch" or Payment_method="Maestro" then %>
								<tr bgcolor='#FFFFFF'>

									<td></td>
									<td><B>Issue Number</B><BR><%= IssueNumber %></td>
									<td><B>Issue Start</B><BR><%= IssueStart %></td>
								</tr>
								<% end if %>
								</TABLE>
								
								
								<% if (maxmind_score="" or isNull(Maxmind_score)) and isNull(masteroid) then %>
								<hr width="600" align="left">
								<table width=600 border=0 cellpadding=5 cellspacing=0>
								<tr bgcolor='#FFFFFF'><td>
								NEW! Protect future transactions with fraud monitoring by <a href=maxmind_settings.asp class=link>Maxmind</a>, <a href=maxmind_settings.asp class=link>click here for details</a>.
								</td></tr>
								</table>
								<% elseif maxmind_score<>"" and isNull(masteroid) then %>
									<hr width="600" align="left">
									<table width=600 border=0 cellpadding=5 cellspacing=0>
									
									<% sql_select = "select * from store_maxmind WITH (NOLOCK) where store_id="&Store_Id&" and oid="&oid
									rs_Store.open sql_select,conn_store,1,1
                                    anonymousProxy=rs_store("anonymousProxy")
									binCountry=rs_store("binCountry")
									binMatch=rs_store("binMatch")
									binName=rs_store("binName")
									binNameMatch=rs_store("binNameMatch")
									binPhone=rs_store("binPhone")
									cityPostalMatch=rs_store("cityPostalMatch")
									shipCityPostalMatch=rs_store("shipCityPostalMatch")
									countryMatch=rs_store("countryMatch")
									countryCode=rs_store("countryCode")
									custPhoneInBillingLoc=rs_store("custPhoneInBillingLoc")
									distance=rs_store("distance")
									sErr=rs_store("err")
									freeMail=rs_store("freeMail")
									highRiskCountry=rs_store("highRiskCountry")
									ip_city=rs_store("ip_city")
									ip_isp=rs_store("ip_isp")
									ip_latitude=rs_store("ip_latitude")
									ip_longitude=rs_store("ip_longitude")
									ip_org=rs_store("ip_org")
									ip_region=rs_store("ip_region")
									maxmindID=rs_store("maxmindID")
									proxyScore=rs_store("proxyScore")
									queriesRemaining=rs_store("queriesRemaining")
									score=rs_store("score")
									spamScore=rs_store("spamScore")                                  
									rs_store.close %>
									<tr bgcolor='#FFFFFF'><td colspan=5>Detailed MaxMind Fraud information below, for definitions on each field <a href=http://www.maxmind.com/app/ccv?promo=STORESEC1452 class=link target=_blank>click here</a></td></tr>
									<% if replace(sErr," ","")<>"" then %>
									<tr bgcolor='#FFFFFF'><TD colspan=5><B>Error:</b><font color=red> <%= sErr %></font></td></tr>
									<% end if %>
                                                                        <tr bgcolor='#FFFFFF'>
									<td><B>Fraud<BR>Score</B><BR><%= score %></td>
									<td><B>Spam<BR>Score</B><BR><%= spamScore %></td>
									<td><B>Proxy<BR>Score</B><BR><%= proxyScore %></td>
									<td><B>Remaining<BR>Queries</B><BR><%= queriesRemaining %></td>
									<td><B>Maxmind<BR>ID</B><BR><%= maxmindID %></td>
									</tr>
									<tr bgcolor='#FFFFFF'>
									<td><B>Country<BR>Code</B><BR><%= countryCode %></td>
									<td><B>Country<BR>Match</B><BR><%= countryMatch %></td>
									<td><B>Postal<BR>Match</B><BR><%= cityPostalMatch %></td>
									<td><B>Phone<BR>Match</B><BR><%= custPhoneInBillingLoc %></td>
									<td><B>High Risk<BR>Country</B><BR><%= highRiskCountry %></td>
									</tr>
									<tr bgcolor='#FFFFFF'>
									<td><B>Free<BR>Mail</B><BR><%= freeMail %></td>
									<td><B>Distance<BR>to IP</B><BR><%= distance %></td>
									<td><B>Anonymous<BR>Proxy</B><BR><%= anonymousProxy %></td>
									<td><B>BankCountry<BR>Match</B><BR><%= binMatch %></td>
									<td><B>Bank<BR>Country</B><BR><%= binCountry %></td>
									</tr>


                           <tr bgcolor='#FFFFFF'>
									<td><B>IP<BR>City*</B><BR><%= ip_city %></td>
									<td><B>IP<BR>Region*</B><BR><%= ip_region %></td>
									<td><B>IP<BR>Longitude*</B><BR><%= ip_longitude %></td>
									<td><B>IP<BR>Latitude*</B><BR><%= ip_latitude %></td>
									<td><B>ISP<BR>Provider</B><BR><%= ip_isp %></td>
									</tr>
									<tr bgcolor='#FFFFFF'><td colspan=4>*Premium fields, only available with paid service</td></tr>

									
									
									</table>

								<% end if %>
								<% if (Processor_id=1 or Processor_id=2 or Processor_id=7 or Processor_id=10 or Processor_id=36 or Processor_id=38) and isNull(masteroid) then %>
									<%
                                                                        if Transaction_Type<>3 then 

%>
										<hr width="600" align="left">
										<table width=600 border=0 cellpadding=5 cellspacing=0>
										<tr bgcolor='#FFFFFF'><td colspan=3>The buttons below will send the transaction back to the payment gateway and update the status as indicated on the selected button.
										This is the same as using your payment gateways virtual terminal to perform the requested change.</td></tr>
										<tr bgcolor='#FFFFFF'><TD align="right" valign=top>
										<% if Transaction_Type = "0" then %>
											<input type="Submit" name="Capture" value="Capture" onClick="return confirmPayment()"><BR>Capture Funds
										<% end if %>
										</td><td align="center">
										<% if (Transaction_Type = "0" or Transaction_Type = "1" or Transaction_Type="2") then %>
											<input type="Submit" name="Credit" value="Credit" onClick="return confirmPayment()"><br>
																		Credit or Capture Amount: <input type=text value="<%= GGrand_Total %>" name="new_amount" size=5>
										<% end if %>
										</td><td align="left" valign=top>
										<% if (Transaction_Type = "0" or Transaction_Type = "1" or Transaction_Type="2") and Processor_id<> 7 then %>
											<input type="Submit" name="Void" value="Void" onClick="return confirmPayment()"><BR>Void Transaction
										<% end if %>
										</td></tr>
										</table>
									<% end if %>

								<% end if %>


						<% End If %>	
					<hr width="600" align="left">

						<table width="600" border="0" cellpadding="2" cellspacing="0">
							
							<tr bgcolor='#FFFFFF'>
							<td valign="top" width="50%"><b>Bill To</b>
									<% Call Display_cust_info(cid,0) %></td>
							<td width="50%" valign="top">
							<% if Show_Shipping then %>
							<b>Ship To</b>
									<blockquote>  
										
										<b><%= ShipCompany %></b><br>
										<%= ShipFirst_name  %> &nbsp;<%= ShipLast_name %><br>
										<%= ShipAddress1 %><br>
										<%=  ShipAddress2 %><br>
										<%=  ShipCity %>, <%= ShipState %>, <%= Shipzip %><br>
										<%= ShipCountry %><br>
										Phone <%= ShipPhone %><br>
										Fax <%= ShipFax %><br>
										<%= ShipEMail %>
									   <% if not(ShipResidential) then %>
										<BR>Residential: No
										<% end if %>
									</blockquote>
							<% end if %>
								</td>
							</tr>
						</table>

						<hr width="600" align="left">
					
						Order Details 
						
						<% text_color="" %>
						<% text_size="2" %>
						<% text_face="Arial" %>
						<% If Request.QueryString("Cart_Type") = "Invoice" then %>
							<a class="link" href="order_details.asp?Oid=<%= Oid %>&Cid=<%= Cid %>&Cart_Type=Big">View Detailed</a>
							<% Call Create_Invoice (Request.QueryString("Cart_Type"),oid) %>
						<% Else	%>
							<a class="link" href="order_details.asp?Oid=<%= Oid %>&Cid=<%= Cid %>&Cart_Type=Invoice">View Invoice</a>
							<% Call Create_Invoice ("Big",oid) %>
						<% End If %>
		
						<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
							
							<tr bgcolor='#FFFFFF'> 
								<TD width="301" height="1" colspan="4"><STRONG>Total</STRONG> </TD>
								<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Total) %></STRONG></TD>
							</tr>
							
							<% if Show_Shipping then %>
							<tr bgcolor='#FFFFFF'>
								<TD width="301" height="1" colspan="4"><STRONG>S &amp; H ( <%= Shipping_Method_Name %>)</STRONG>  </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Shipping_Method_Price )%></STRONG></TD>
							</tr>
							<% end if %>
							<tr bgcolor='#FFFFFF'> 
								<TD width="301" height="1" colspan="4"><STRONG> Tax </STRONG></TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Tax) %></STRONG></TD>
							</tr>
							
							<%  If Coupon_Id <> "" or Coupon_Amount>0 then	  %>
								<tr bgcolor='#FFFFFF'>
									<td align="left" colspan="4"><STRONG> <i><%= Coupon_Id %></i> </STRONG> &nbsp;</td>
									<td align="left"><STRONG><i>-<%= Currency_Format_Function(Coupon_Amount) %></i></STRONG></td>
								</tr>
							<% End If %>
							<% If cdbl(rewards_Amount) > cdbl(0) then %>
							 	<tr bgcolor='#FFFFFF'>
									<td align="left" colspan="4"><STRONG> <i>Rewards</i> </STRONG> </td>
									<td align="left"><STRONG><i>-<%= Currency_Format_Function(rewards_Amount) %></i></STRONG></td>
								</tr>
							<% end if %>
							<% If cdbl(Gift_Wrapping) <> 0 then %>
							 	<tr bgcolor='#FFFFFF'>
									<td align="left" colspan="4"><STRONG> Gift Wrapping </STRONG> </td>
									<td align="left"><STRONG><i><%= Currency_Format_Function(Gift_Wrapping_amount) %></i></STRONG>&nbsp;</td>
								</tr>
							<% end if %>
							<% If Giftcert_Id <> "" and cdbl(Giftcert_Amount)>0 then %>
							 	<tr bgcolor='#FFFFFF'>
									<td align="left" colspan="4"><STRONG>Gift Certificate </STRONG> </td>
									<td align="left"><STRONG><i>-<%= Currency_Format_Function(Giftcert_Amount) %></i></STRONG></td>
								</tr>
							<% end if %>


							<tr bgcolor='#FFFFFF'>
								<TD width="304" height="1" colspan="4"><STRONG> Grand total</STRONG> </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Grand_Total) %></STRONG></TD>
						</tr>
						</TABLE>

						<% if Recurring_Total <> 0 then %>
						<BR><BR>
						<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
							
							<tr bgcolor='#FFFFFF'> 
								<TD width="301" height="1" colspan="4"><STRONG>Recurring Total</STRONG> </TD>
								<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Recurring_FeeT) %></STRONG></TD>
							</tr>

							<tr bgcolor='#FFFFFF'>
								<TD width="301" height="1" colspan="4"><STRONG>Recurring Tax</STRONG></TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Recurring_Tax) %></STRONG></TD>
							</tr>
							<tr bgcolor='#FFFFFF'>
								<TD width="304" height="1" colspan="4"><STRONG>Every <%= Recurring_Days %> days</STRONG> </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Recurring_Total) %></STRONG></TD>
						</tr>
						</TABLE>
						<% end if %>
					</td>
				</tr>
			</table>

			<hr width="600" align="left">
			Total Order Weight=<%= Session("Total_Weight") %><BR><BR>
			<% if Referer <> "" then %>
				Referring Url=<a href="<%=Referer %>" target=_blank><%= Referer %></a>
				<BR><BR><% end if %>
				<% if ClientIp <> "" then %>
				Client IP Address=<%= ClientIp %>
				<BR><BR>
				<% end if %>
				<% ' ORDER NOTES

			set rsOrderNotes = server.createobject("ADODB.Recordset")
			sql_ordernotes = "select sys_created, notes from store_purchases_notes WITH (NOLOCK) where  store_id=" &store_id&" and order_id = "&Oid&" order by sys_created desc"
			rsOrderNotes.open sql_ordernotes,conn_store,1,1
		    if not rsOrdersNotes.eof then
			%>
			Order Notes
			<table cellspacing="0" cellpadding="0" border="1" width='600'>
				<% while not rsOrderNotes.eof %>	
				<tr bgcolor='#FFFFFF'>
					<td width="5%" nowrap><%=FormatDateTime(rsOrderNotes("sys_created"))%>&nbsp;</td>
					<td width="95%">&nbsp;&nbsp;<%=rsOrderNotes("notes")%></td>
				</tr>

			<% 
			rsOrderNotes.movenext
			wend
			%>
						</table>
			<% end if %>

			<form method="POST" action="" id=form1 name=form1>
			<table border="0" width="77%">
			     
				<% If Is_Gift = -1 and (Gift_Message <> "" or Gift_Wrapping = "Yes") then %>
					<tr bgcolor='#FFFFFF'>
					<td width="25%"><b>Gift?</b> <%= Is_Gift_Text %> </td>
					<td></td>
					</tr>
				<% End If %>

				<% if Gift_Message <> "" then %>
				<tr bgcolor='#FFFFFF'><td width="102%" colspan="3"><b>Gift Message</b>
        <BR><%= Gift_Message %></td></tr>
				<% end if %>

				<% if Cust_Notes <> "" then %>
				<tr bgcolor='#FFFFFF'>
				<td width="102%" colspan="3"><b>Customer Notes</b><br>
					<%= Cust_Notes %></td>
				</tr>
				<% end if %>
				<% if Custom_fields <> "" then %>
				<tr bgcolor='#FFFFFF'>
				<td width="102%" colspan="3"><b>Custom Fields</b><br>
					<%= Custom_fields %></td>
				</tr>
				<% end if %>
				<tr bgcolor='#FFFFFF'>
				<td width="102%" colspan="3">
				<hr width="600" align="left"><b>
						<b>
						<a class="link" href="ship_orders.asp?oid=<%= oid %>&cid=<%= cid %>">Ship Order</a></b>&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;
						<b><a class="link" href="return_orders.asp?oid=<%= oid %>&cid=<%= cid %>">Return Order</a></b>
				      &nbsp;&nbsp;&nbsp;&nbsp;
					  <b><a href="Javascript:fnAddNotes('<%=Oid%>')" class=link>Add Notes</a></b>
					  &nbsp;&nbsp;&nbsp;&nbsp;
					  <b><a class="link" href="orders.asp?delete_id=<%= oid %>">Delete Order</a></b></td>

            </tr>

			</table>
			</form>


				
				<%
				sql_select_shippments = "Select * from Store_Purchases_shippments WITH (NOLOCK) where oid = "&Oid&" and Store_id ="&Store_id&" order by sys_created desc"
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_select_shippments,mydata,myfields,noRecords)

				if noRecords = 0 then %>
				<table border="1" cellspacing="0">

				<tr bgcolor='#FFFFFF'>
				<td><b>Shipped</b></td>
				<td><b>Company</b></td>
				<td><b>Date Shipped</b></td>
				<td><b>Tracking ID</b></td>
				<td><b>Shipping Notes</b></td>
			</tr>
				<% FOR rowcounter= 0 TO myfields("rowcount") %>

				<tr bgcolor='#FFFFFF'>
					<td align="middle"><%= mydata(myfields("shippedpr"),rowcounter) %>&nbsp;</td>
					<td align="middle"><%= mydata(myfields("shipping_company"),rowcounter) %>&nbsp;</td>
					<td align="middle"><%= FormatDateTime(mydata(myfields("shipped_date"),rowcounter)) %>&nbsp;</td>
					<td align="middle"><%= mydata(myfields("tracking_id"),rowcounter) %>&nbsp;</td>
					<td align="middle"><%= mydata(myfields("shippment_notes"),rowcounter)%>&nbsp;</td>
				</tr>
				<% Next %>
				</table>
				<% End If %>

				<BR><BR>

				

				<% sql_select_returns = "Select * from Store_Purchases_Returns WITH (NOLOCK) where oid = "&Oid&" and Store_id ="&Store_id
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_select_returns,mydata,myfields,noRecords)

				if noRecords = 0 then %>
				<table border="1" cellspacing="0">
				<tr bgcolor='#FFFFFF'>
				<td><b>Return Requested</b></td>
				<td><b>RMA Number</b></td>
				<td><b>Return Received</b></td>
				<td><b>Return Notes</b></td>
			</tr>
				<% FOR rowcounter= 0 TO myfields("rowcount") %>
				<tr bgcolor='#FFFFFF'>
					<td align="middle"><%= mydata(myfields("sys_created"),rowcounter) %>&nbsp;</td>
					<td><%= mydata(myfields("return_rma"),rowcounter) %>&nbsp;</td>
					<td align="middle"><%= FormatDateTime(mydata(myfields("return_date"),rowcounter)) %>&nbsp;</td>
					<td align="middle"><%= mydata(myfields("return_notes"),rowcounter)%>&nbsp;</td>
				</tr>
				<% Next %>
				</table>
				<% End If %>

				


		</td>
	</tr>
<% createFoot thisRedirect, 0%>
