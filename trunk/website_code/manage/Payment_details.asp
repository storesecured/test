<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include virtual="common/crypt.asp"-->
<%
Payment_Id = Request.QueryString("Payment_Id")
store_id1 = Request.QueryString("store_id1")
if not isNumeric(Soid) then
	Response.Redirect "error.asp?message_id=1"
end if

on error resume next
sql_select = "select * from Sys_Billing where Store_Id = "&Store_Id1

rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	First_name= rs_store("First_name")
	Last_name= rs_store("Last_name")
	Company= rs_store("Company")
	Address= rs_store("Address")
	City= rs_store("City")
	State= rs_store("State")
	Zip= rs_store("Zip")
	Country= rs_store("Country")
	Phone= rs_store("Phone")
	Fax= rs_store("Fax")
	EMail= rs_store("EMail")
	cardnumber =  rs_store("card_Number")
	exp_month=rs_store("exp_month")
	exp_year=rs_store("exp_year")

end if
rs_store.close

sql_select = "select * from sys_payments where store_id="&Store_Id1&" and Payment_Id="&Payment_Id
rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	Shipping_Method_Price = 0
	Shipping_Method_Name = "Service"
	Tax = 0
	Grand_Total = Rs_store("Amount")
	Purchase_Date = Rs_store("TransDate")
	Payment_Description = rs_store("Payment_Description")
	Card_Ending = rs_store("Card_Ending")
	Auth_number = rs_Store("Auth_number")
        Transaction_Id = rs_Store("Transaction_Id")


	AVS_result = rs_store("AVS_result")
	Card_verif = rs_store("Card_verif")
	IP_Address = rs_store("IP_Address")
	IP_Country = rs_store("IP_Country")
	Fraud_score = rs_store("Fraud_score")
	fraud_id = rs_store("fraud_id")
	
	if isNull(Auth_Number)  then 
		Auth_Number ="Unknown"
	end if
	if isNull(AVS_result) then 
		AVS_result ="Unknown"
	end if
	if isNull(Card_verif) then 
		Card_verif ="Unknown"
	end if
	if isNull(IP_Address) then 
		IP_Address ="Unknown"
	end if
	if isNull(IP_Country) then 
		IP_Country ="Unknown"
	end if
	if isNull(Fraud_score)  then 
		Fraud_score ="Unknown"
	end if
	if isNull(fraud_id) then 
		fraud_id ="Unknown"
	end if


end if
rs_store.close

if isNumeric(Card_Ending) then
	Payment_Method = "****" & Card_Ending
else
   Payment_Method = Card_Ending
end if

sFormAction = "payment_details.asp"
sSubmitName = "update_promotion_ids"
sTitle = "Payment Details"
thisRedirect = "payment_details.asp"
sMenu = "account"
createHead thisRedirect



%>
<!--BEGIN REFUND PAYMENT CODE-->
<%
if Request.Form("Credit") <> ""  then
		
		GGrand_Total = Request.Form("new_amount")
                Orig_id=request.form("trans_id")
		%>
		<!--#include file="include\verisign\verisign_refund.asp"-->
		<%


end if
%>
<!--END REFUND PAYMENT CODE-->
<tr><td>
<table width=100%>
				<tr bgcolor='#FFFFFF'>
					<td><b>Store ID <%=Store_Id1%><input type="hidden" name="store_id" value="<%=Store_Id1%>"></b></td>
				</tr>

	<tr bgcolor='#FFFFFF'>
		<td>
			<table width=100% border=0 cellpadding=2 cellspacing=0>

				<tr bgcolor='#FFFFFF'>
					<td valign=top width=50% class='normal'>
						<B>Current Billing Details</B><BR>
						<% if Company <> "" then %>
						<%= Company %><BR>
						<% end if %>
                  <%= First_name %>&nbsp;<%= Last_Name %><BR>
                  <%= Address %><BR>
                  <%= City %>,&nbsp;<%= State %>&nbsp;<%= Zip %><BR>
                  <%= Country %><BR>
                  <%= Phone %><BR>
                  <%= Email %><BR>
				
				  <%
				  If strServerPort<> 443 then
						if CardNumber<>"" then
							cardnumber1=cardnumber
							BeginCardNumber = Left(DeCrypt(CardNumber),4)
							EndCardNumber = Right(DeCrypt(CardNumber),4)
							CardNumber = BeginCardNumber & "**********" & EndCardNumber
							CardNumber = CardNumber&"<font size=1><BR><a href=https://"&request.servervariables("HTTP_HOST")&"/payment_details.asp?store_id1="+ Store_Id1 + "&Payment_Id="+Payment_id+">Switch to Secure mode to see the entire number</a></font>"
						else
							CardNumber = "Unknown"
						end if
					else
						CardNumber = DeCrypt(CardNumber)
					End if
					response.write "<b>CC# &nbsp;&nbsp;&nbsp;" + cardNumber +"</b>"
	%>
					</TD>


				</tr>
			</table>	
			
			<hr width=100% align=left>


	
				<table width=100% border=0 cellpadding=2 cellspacing=0>
				<tr bgcolor='#FFFFFF'>
					<td><b>Payment Details</b></td>
				</tr>


				<tr bgcolor='#FFFFFF'>
					<td><b>Payment ID <%=Payment_id%></b></td>
				</tr>
					<td><b>Auth Number <%=Auth_number%></b></td>
				<tr bgcolor='#FFFFFF'>
					
					<td class='normal'><B>Date </B><%= FormatDateTime(Purchase_Date,2) %></td>
					<td class='normal'><B>Order ID </B><%= Store_Id1 %>-<%= Payment_ID %></td>
					<td class='normal'><B>Payment </B><%= Payment_Method %></td>

				</tr>
			</table>

			<hr width=100% align=left>
			
			<font CLASS='normal'>Bill Details</font>
         <table cellspacing="0" cellpadding="0" border="1"  width="100%" >
			<tr bgcolor='#FFFFFF'>   

				<td class='normal'><STRONG>Quantity</STRONG></td>
				<td class='normal' width=30%><STRONG>Name</STRONG></td>
				<td class='normal'><STRONG>Price</STRONG></td>
				<td class='normal'><STRONG>Total</STRONG> </td>

			</tr>

			<tr bgcolor='#FFFFFF'>
				<td class='normal'>1</td>
	    		<td class='normal'><%= Payment_Description %></td>
				<td class='normal'>$<%= FormatNumber(Grand_Total,2) %></td>
    			<td class='normal'>$<%= FormatNumber(Grand_Total,2) %></td>
			</tr>

		</table>
	
<br>


		<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
			<tr bgcolor='#FFFFFF'>
				<TD width="201" height="1" colspan="4" class='normal'><STRONG>Total</STRONG> </TD>
				<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Grand_Total,2) %></STRONG></TD>
			</tr>
				<tr bgcolor='#FFFFFF'>
					<TD width="201" height="1" colspan="4" class='normal'><STRONG>Shipping and Handling</STRONG>  </TD>
				<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Shipping_Method_Price,2) %></STRONG></TD>
			</tr>
				<tr bgcolor='#FFFFFF'>
					<TD width="201" height="1" colspan="4" class='normal'><STRONG>Tax </STRONG></TD>
					<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Tax,2) %></STRONG></TD>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<TD width="204" height="1" colspan="4" class='normal'><STRONG>Grand Total</STRONG> </TD>
					<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Grand_Total,2) %></STRONG></TD>
			</tr>
			</TABLE>
		</td>
	</tr>

	<tr bgcolor='#FFFFFF'>
	<td>
				<hr width=100% align=left>
		<font color="#336699"><b>Card Authorization & Verification Details</b></font>
		<table width="100%" border="1" cellspacing="0" cellpadding="0">
			<tr bgcolor='#FFFFFF'>
				<td ><b>Authorization Number</b></td>
				<td Colspan=3><input name="AuthNumber" value=<%=Auth_Number%>></td>

			</tr>
			<tr bgcolor='#FFFFFF'>
				<td ><b>Transaction Id</b></td>
				<td Colspan=3><input name="Transaction_Id" value=<%=Transaction_Id%>></td>

			</tr>


			<tr bgcolor='#FFFFFF'>
				<td ><b>cardnumber</b></td>
				<td Colspan=3><input name="cardnumber" value=<%=CardNumber1%>></td>


			</tr>
			<tr bgcolor='#FFFFFF'>
				<td ><b>Expiry</b></td>
				<td Colspan=3><input name="exp" value=<%=trim(cstr(exp_month))&trim(cstr(exp_year))%>></td>


			</tr>

			<tr bgcolor='#FFFFFF'>
				<td><b>AVS Result</b></td>
				<td Colspan=3><%=AVS_result%></td>

			</tr>

			<tr bgcolor='#FFFFFF'>
				<td><b>Card Verification</b></td>
				<td Colspan=3><%=Card_verif%></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
				<td><b>IP Address</b></td>
				<td>	<%=IP_Address%></td>
				<td><b>IP Country</b></td>
				<td>	<%=IP_Country%></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
				<td><b>Fraud Score</b></td>
				<td><%=Fraud_score%></td>
				<td><b>Fraud ID</b></td>
				<td><%=fraud_id%></td>

			</tr>
		</table>

	</td>
	</tr>

</table>
</td></tr>

<!--BEGIN REFUND PAYMENTS BUTTONS-->
<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
function confirmPayment()
{
var agree=confirm("Are you sure?");
if (agree)
	return true ;
else
	return false ;
}
</script>
<% if Payment_Method <> "Paypal" then %>
                                <table width=100% border=0 cellpadding=5 cellspacing=0>
								<tr bgcolor='#FFFFFF'>
								<td colspan=2>The buttons below will send the transaction back to the payment gateway and update the status as indicated on the selected button. This is the same as using your payment gateways virtual terminal to perform the requested change.
								</td>
								</tr>
								<tr bgcolor='#FFFFFF'>
								<td>
								<input type=hidden name=Card_Ending value='<%=Card_Ending %>'>
								<input type=hidden name=trans_id value='<%=Transaction_Id %>'>
								Description: <input type=text name="refundDescription" Value="Amount Refunded" size="30">
								<td align="center">
								Credit Amount: <input type=text value="<%= Grand_Total %>" name="new_amount" size=5>
								<% 'if (Transaction_Type = "0" or Transaction_Type = "1" or Transaction_Type="2") then %>
								<input type="Submit" name="Credit" value="Credit" onClick="return confirmPayment()">&nbsp;&nbsp;
                                                                
								<% 'end if %>
								</tr>
								</table>
<% else %>
</form>
You must already be logged into paypal
<form method="post" action="https://history.paypal.com/us/cgi-bin/webscr" target=_Blank>
<input type="hidden" id="" name="return_to" value="">
<input type="hidden" id="" name="history_cache" value="">
<input type="hidden" id="" name="item" value="0">
<input type="hidden" name="search_type" value="<%=Auth_Number%>">
<input type=hidden name="for" value=4>
<input type="hidden" name="span" value="broad">
<input type="hidden" id="" name="cmd" value="_history-search">
<input type="submit" name="submit.x" value="Search for Transaction in Paypal">
<input type="hidden" name="form_charset" value="UTF-8">
<input type=hidden name="search_first_type" value="trans_id">
</form>

<% end if %>

				
<!--END OF REFUND PAYMENTS BUTTONS-->

<% createFoot thisRedirect, 0%>

