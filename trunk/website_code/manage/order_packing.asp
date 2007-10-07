<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<title>StoreSecured Packing Slip - #<%= Request.QueryString("Oid") %></title>
<style>td {COLOR: 000000;TEXT-DECORATION: none;font-size : 10pt;font-family :Verdana;}
td.to {COLOR: 000000;TEXT-DECORATION: none;font-size : 10pt;font-family :Verdana;}</style>
<base href="<%= Site_Name %>">
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">

<%

' Custom Invoice Header & Footer
sql_Select_Invoice_Custom = "Select Invoice_header, Invoice_footer From Store_Settings where Store_id = "&Store_id
rs_Store.open sql_Select_Invoice_Custom,conn_store,1,1
rs_Store.MoveFirst 
	Invoice_header = Rs_store("Invoice_header")
	Invoice_footer = Rs_store("Invoice_footer")
rs_Store.Close

'LOAD ORDER DETAILS FOR SPECIFIED ORDER FROM THE DATABASE
Oid = Request.QueryString("Oid")
Cid= Request.QueryString("Cid")
if cid="" or oid="" then
   response.redirect "error.asp?Message_Id=1"
end if

sql_select_ccid = "select ccid from store_customers where record_type=0 and cid="&cid
rs_store.open sql_select_ccid, conn_Store, 1, 1
if not rs_Store.bof and not rs_Store.eof then
   ccid = rs_Store("CCID")
else
  ccid=0
   'response.redirect "error.asp?Message_Id=1"
end if
rs_store.close


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
	Coupon_id = Rs_store("Coupon_id")
	Coupon_Amount = Rs_store("Coupon_Amount")
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
	 ShipPhone	=Rs_store("ShipPhone")
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
rs_Store.Close

 %>


<table width=600 border=0 cellpadding=2 cellspacing=0>
<tr> 
<td class='normal'><%=Invoice_header%></td>
</tr>




	<tr>
		<td width="100%" colspan="3" height="74">
			<table width="600">
				<tr valign=top>
				<td colspan=2>
						<table width="600" border="0" cellpadding="2" cellspacing="0">
                                                        <tr><td>
                                                        <%= Store_Company %><br>
                                                	<%= Store_Address1 %>
										   <% if Store_Address2<>"" then %>
										   <BR><%= Store_Address2 %>
										   <% end if %>
										   <br><%= Store_City %>, <%= Store_State %>, <%= Store_Country %>, <%= Store_Zip %> <br> Phone: <%= Store_Phone %>, Fax: <%= Store_Fax %><br>Email: <%= Store_Email %><br>
                                                	</td>
                                                        <td valign=top align=right>Order #: <%= oid %></td></tr>
                                                	<tr><td colspan=2><HR></td></tr>

							<tr>
							<td valign="top" width="50%"><b>Bill To</b>
									<% Call Display_cust_info(cid,0) %></td>
							<td width="50%" valign="top">
							<% 
							if Show_Shipping then %>
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
						<% Call Create_Invoice ("Packing",oid) %>

		
						<% if show_shipping then%>
						<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
							

							
							<tr>
								<TD width="301" height="1" colspan="4"><STRONG>S &amp; H ( <%= Shipping_Method_Name %>)</STRONG>  </TD>
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
					<td width="25%"><b>Wrap Order</b> <%= Gift_Wrapping %></td>
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