<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->

<html><head>
<title>StoreSecured Printable Order - #<%= Request.QueryString("Oid") %></title>
<style>td {COLOR: 000000;TEXT-DECORATION: none;font-size : 10pt;font-family :Verdana;}
td.to {COLOR: 000000;TEXT-DECORATION: none;font-size : 10pt;font-family :Verdana;}</style>
<base href="<%= Site_Name %>">
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<%

'LOAD ORDER DETAILS FOR SPECIFIED ORDER FROM THE DATABASE
Oid = Request.QueryString("Oid")
Cid= Request.QueryString("Cid")
if cid="" or oid="" then
   response.redirect "error.asp?Message_Id=1"
end if

' Custom Invoice Header & Footer
sql_Select_Invoice_Custom = "Select Invoice_header, Invoice_footer From Store_settings where Store_id = "&Store_id
rs_Store.open sql_Select_Invoice_Custom,conn_store,1,1
rs_Store.MoveFirst 
	Invoice_header = Rs_store("Invoice_header")
	Invoice_footer = Rs_store("Invoice_footer")
rs_Store.Close


sql_select_ccid = "select ccid from store_customers where record_type=0 and cid="&cid
rs_store.open sql_select_ccid, conn_Store, 1, 1
if not rs_Store.bof and not rs_Store.eof then
   ccid = rs_Store("CCID")
else
   cci=0
end if
rs_store.close

if Request.Form("Capture") <> "" or Request.Form("Credit") <> "" then
	  if Request.Form("Capture") <> "" then
	sType = "Capture"
	trans_type = "2"
	  else
	sType = "Credit"
	trans_type = "3"
	  end if
	  Real_Time_Processor = Request.form("Processor_id")
	  GGrand_Total = Request.Form("new_amount")
	if Real_Time_Processor = 2 then
		%><!--#include file="include\Authorizenet\Authorizenet.asp"--><%
	elseif Real_Time_Processor = 7 then
	  %><!--#include file="include\Echo\Echo.asp"--><%
	  elseif Real_Time_Processor = 9 then
	  %><!--#include file="include\Verisign\verisign.asp"--><%
	  end if
	sql_update_purchases = "update store_purchases set Transaction_Type='"&trans_type&"' where (oid = "&oid&" or masteroid="&oid&") AND Store_id ="&Store_id
	conn_store.Execute sql_update_purchases
end if

'LOAD VALUES FROM STORE_PURCHASES TABLE
sql_select_purchases = "Select * from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id
rs_Store.open sql_select_purchases,conn_store,1,1
if rs_store.bof and rs_store.eof then
   response.redirect "error.asp?Message_Id=1"
end if
	Verified = Rs_store("Verified")
	if Verified = -1 then
		checked = "checked"
	end if
	Verified_Ref = Rs_store("Verified_Ref")
	Invoice_Id = Rs_Store("Invoice_Id")
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Shipping_Method_Name = Rs_store("Shipping_Method_Name")
	Tax = Rs_store("Tax")
	Total = Rs_store("Total")
	Coupon_id = Rs_store("coupon_id")
     Coupon_Amount = Rs_store("coupon_amount")
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
	Verified_Ref = Rs_store("Verified_Ref")
	AVS_Result = Rs_store("AVS_Result")
	Card_Code_Verification = Rs_store("Card_Code_Verification")
	Processor_id = Rs_store("Processor_id")
	Transaction_Type = Rs_store("Transaction_Type")
	Purchase_Date = Rs_store("Purchase_Date")
	Cust_Po = Rs_store("Cust_PO")
	Cust_Notes = Rs_store("Cust_Notes")
	Custom_fields = Rs_store("Custom_fields")
	Is_Gift = Rs_store("Is_Gift")
	Gift_Wrapping = Rs_store("Gift_Wrapping")
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
	 ShipPhone	=Rs_store("ShipPhone")
	 ShipFax =Rs_store("ShipFax")
	 ShipEMail=Rs_store("ShipEMail")
	 ShipResidential=rs_store("ShipResidential")
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
rs_Store.Close

on error resume next

if Request.QueryString("Cart_Type") = "Invoice" then
	sTitle = "Invoice"
else
	sTitle = "Order"
end if
 %>

<table width=600 border=0 cellpadding=2 cellspacing=0>
<tr>
<td class='normal'><%=Invoice_header%></td>
</tr>
<% if Store_id=15340 then %>
<tr>
<td class='normal' align=center>
<hr width="600" align="left">
<font size=5><%= sTitle %> # <%= oid %></font>
<hr width="600" align="left">
</td>
</tr>
<% end if %>

	<tr>
		<td width="600" colspan="3" height="74">
			<table width="600">
				

				<tr valign=top>
				<td colspan=2>
						<table width=600 border=0 cellpadding=2 cellspacing=0>

							<tr>
							<td><B>Date:</B>  <%= FormatDateTime(Purchase_Date,2) %></td>
							<td><B>Status:</B> <%= Verified %></td>
							<td><B>Customer ID:</B> <%= CCid %></td>
							</tr>
							
							<tr>
							<td><B><%= sTitle %> ID:</B> <%= Oid %></td>
								<td><B>Terms:</B> <%= Payment_Method %></td>
								<td><b>Purchase Order:</b> <%= Cust_PO %></td>
							</tr>

						</table>
						
						<% if Payment_Method = "Visa" or Payment_Method = "Mastercard" or Payment_Method = "Discover" or Payment_Method = "American Express" or Payment_Method = "Diner's Club" then
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
								<tr>
									<td rowspan=2 width=15%><B>Credit card order,</b> processed by <%= Processor %></td>
									<td width=25%><B>Card Number</B><BR><%= CardNumber %></td>
									<td width=25%><B>Card Name</B><BR><%= CardName %></td>
									<td width=10%><B>Auth #</B><BR><%= left(AuthNumber,10) %></td>
									<td width=25%><b>Card Exp</b><BR><%= CardExpiration %></td>

								</tr>
								<tr>

									<td><B>AVS Result</B><BR><%= AVS %></td>
									<td><B>Code Result</B><BR><%= Card_Code %></td>
									<td><B>Ref #</B><BR><%= left(Verified_Ref,10) %></td>
									<td><b>Type</b><BR><%= Trans_Type %></td>
						 <input type=hidden name="Processor_id" value="<%= Processor_id %>">
								</tr>
								<% if strServerPort= 443 then %>
								<TR><TD colspan=2 align="right">
								<% if Transaction_Type = "0" and 1=0 then %>
								<input type="Submit" name="Capture" value="Capture"><BR><font size=1>Capture Funds from Previous Authorization</font>
								<% end if %>
								</td><td align="right">
								<% if (Transaction_Type = "0" or Transaction_Type = "1") and 1=0 then %>
								<input type=text value="<%= Grand_Total %>" name="new_amount"></td><td colspan=2 align=left>&nbsp;&nbsp;<input type="Submit" name="Credit" value="Credit"><BR><font size=1>Credit Funds back to Customer</font>
								<% end if %>
								</td></tr>
								<% end if %>
							</table>
						<% End If %>	
					<hr width="600" align="left">
						<table width="600" border="0" cellpadding="2" cellspacing="0">
							
							<tr>
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
							<%end if%>
								</td>
							</tr>
						</table>

						<hr width="600" align="left">
					
						Order Details 
						
						<% text_color="" %>
						<% text_size="2" %>
						<% text_face="Arial" %>
						<% If Request.QueryString("Cart_Type") = "Invoice" then %>	  
							<% Call Create_Invoice (Request.QueryString("Cart_Type"),oid) %>
						<% Else	%>
							<% Call Create_Invoice ("Big",oid) %>
						<% End If %>
		
						<table cellspacing="0" cellpadding="0" border="1"	width="600" >
							
							<tr> 
								<TD width="301" height="1" colspan="4"><STRONG>Total</STRONG> </TD>
								<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Total) %></STRONG></TD>
							</tr>
							
							<% if Show_Shipping then%>
							<tr>
								<TD width="301" height="1" colspan="4"><STRONG>S &amp; H ( <%= Shipping_Method_Name %>)</STRONG>  </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Shipping_Method_Price )%></STRONG></TD>
							</tr>
							<% end if %>
							<% if tax<>0 then %>
                                                        <tr>
								<TD width="301" height="1" colspan="4"><STRONG> Tax </STRONG></TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Tax) %></STRONG></TD>
							</tr>
							<% end if %>
							
							
							<%  If Coupon_Id <> "" or Coupon_Amount>0 then	  %>
								<tr bgcolor='#FFFFFF'>
									<td align="left" colspan="4"><STRONG> <i><%= Coupon_Id %></i> </STRONG> </td>
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
									<td align="left"><STRONG><i><%= Currency_Format_Function(Gift_Wrapping_amount) %></i></STRONG></td>
								</tr>
							<% end if %>
							<% If Giftcert_Id <> "" and cdbl(Giftcert_Amount)>0 then %>
							 	<tr bgcolor='#FFFFFF'>
									<td align="left" colspan="4"><STRONG>Gift Certificate </STRONG> </td>
									<td align="left"><STRONG><i>-<%= Currency_Format_Function(Giftcert_Amount) %></i></STRONG></td>
								</tr>
							<% end if %>

							<tr>
								<TD width="304" height="1" colspan="4"><STRONG> Grand total</STRONG> </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Grand_Total) %></STRONG></TD>
						</tr>
						</TABLE>

						<% if Recurring_Total <> 0 then %>
						<BR><BR>
						<table cellspacing="0" cellpadding="0" border="1"	width="600" >
							
							<tr> 
								<TD width="301" height="1" colspan="4"><STRONG>Recurring Total</STRONG> </TD>
								<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Recurring_FeeT) %></STRONG></TD>
							</tr>

							<tr>
								<TD width="301" height="1" colspan="4"><STRONG>Recurring Tax </STRONG></TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Recurring_Tax) %></STRONG></TD>
							</tr>

							<tr>
								<TD width="304" height="1" colspan="4"><STRONG>Every <%= Recurring_Days %> days</STRONG> </TD>
							<TD width="71" height="1"><STRONG><%= Currency_Format_Function(Recurring_Total) %></STRONG></TD>
						</tr>
						</TABLE>
						<% end if %>
					</td>
				</tr>
			</table>

			<table border="0" width="77%">
			
				<% If Is_Gift = -1 and (Gift_Message <> "" or Gift_Wrapping = "Yes") then %>
				
					<tr>
					<td width="25%"><b>Gift?</b> <%= Is_Gift_Text %> </td>
					<td></td>
					</tr>
				<% End If %>

				<% if Gift_Message <> "" then %>
				<TR><td width="102%" colspan="3"><b>Gift Message</b><BR>
        <%= Gift_Message %></td></tr>
				<% end if %>

				<% if Cust_Notes <> "" then %>
				<tr>
				<td width="102%" colspan="3"><b>Customer Notes</b><br>
					<%= Cust_Notes %>
				</tr>
				<% end if %>
				<% if Custom_fields <> "" then %>
				<tr>
				<td width="102%" colspan="3"><b>Custom Fields</b><br>
					<%= Custom_fields %></td>
				</tr>
				<% end if %>


			</table>


				
				<%
				sql_select_shippments = "Select * from Store_Purchases_shippments where oid = "&Oid&" and Store_id ="&Store_id
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_select_shippments,mydata,myfields,noRecords)

				if noRecords = 0 then %>
				<BR><BR><table border="1" cellspacing="0" width=600>
				
				<tr>
				<td><b>Shipped</b></td>
				<td><b>Shipping Method</b></td>
				<td><b>Shipping Costs</b></td>
				<td><b>Date Shipped</b></td>
				<td><b>Tracking ID</b></td>
				<td><b>Shipping Notes</b></td>
			</tr>
				
				<% FOR rowcounter= 0 TO myfields("rowcount") %>

				<tr>
					<td align="middle"><%= mydata(myfields("shippedpr"),rowcounter) %></td>
					<td><%= Shipping_Method_Name %></td>
					<td align="middle"><%= Currency_Format_Function(Shipping_Method_Price) %></td>
					<td align="middle"><%= FormatDateTime(mydata(myfields("shipped_date"),rowcounter)) %></td>
					<td align="middle"><%= mydata(myfields("tracking_id"),rowcounter) %></td>
					<td align="middle"><%= mydata(myfields("shippment_notes"),rowcounter)%></td>
				</tr>
				<% Next %>
				</table>
				<% End If %>



				<% sql_select_returns = "Select * from Store_Purchases_Returns where oid = "&Oid&" and Store_id ="&Store_id
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_select_returns,mydata,myfields,noRecords)

				if noRecords = 0 then %>
				<BR><BR>
				<table border="1" cellspacing="0">
				
			 <tr>
				<td><b>Return Requested</b></td>
				<td><b>RMA Number</b></td>
				<td><b>Return Received</b></td>
				<td><b>Return Notes</b></td>
			</tr>
				<% FOR rowcounter= 0 TO myfields("rowcount") %>
				<tr>
					<td align="middle"><%= mydata(myfields("sys_created"),rowcounter) %></td>
					<td><%= mydata(myfields("return_rma"),rowcounter) %></td>
					<td align="middle"><%= FormatDateTime(mydata(myfields("return_date"),rowcounter)) %></td>
					<td align="middle"><%= mydata(myfields("return_notes"),rowcounter)%></td>
				</tr>
				<% Next %>
			</table>
				<% End If %>

		</td>
	</tr>
	
	<tr>
		<td width="600" colspan="3" height="74">
			<table width=600 border=0 cellpadding=2 cellspacing=0>
			<tr> 
			<td class='normal'><%=Invoice_footer%></td>
			</tr>
			</table>
		</td>
	</tr>
