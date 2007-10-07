<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sFormAction = "my_invoice.asp"
sSubmitName = "update_promotion_ids"
sTitle = "Prior Invoice Detail"
sFullTitle = "My Accounts > Payments > <a href=my_payments.asp class=white>Prior Invoices</a> > Detail"
thisRedirect = "my_invoice.asp"
sMenu = "account"
createHead thisRedirect

Payment_Id = Request.QueryString("Payment_Id")
if not isNumeric(Soid) then
	Response.Redirect "error.asp?message_id=1"
end if

on error resume next
sql_select = "select * from Sys_Billing where Store_Id = "&Store_Id

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
else

	Company= Store_Company
	Address= Store_Address1
	City= Store_City
	Zip= Store_Zip
	Country= Store_Country
	Phone= Store_Phone
	Fax= Store_Fax
	EMail= Store_Email
end if
rs_store.close

sql_select = "select * from sys_payments where store_id="&Store_Id&" and Payment_Id="&Payment_Id
rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	Shipping_Method_Price = 0
	Shipping_Method_Name = "Service"
	Tax = 0
	Grand_Total = Rs_store("Amount")
	Purchase_Date = Rs_store("TransDate")
	Payment_Description = rs_store("Payment_Description")
	Card_Ending = rs_store("Card_Ending")
end if
rs_store.close

if isNumeric(Card_Ending) then
	Payment_Method = "****" & Card_Ending
else
   Payment_Method = Card_Ending
end if

%>


	<tr bgcolor=FFFFFF>
		<td>
			<hr size=1 noshade>

			<font CLASS='big'>Invoice</FONT><BR>
			<font CLASS='normal'><b>StoreSecured</b></font><BR>
		<font CLASS='normal'>10272 Aviary Drive&nbsp;
			<BR>San Diego, CA, USA, 92131<BR>
			866-324-2764<BR>
				<%=sSupport_email %><BR></font>

			<hr>
			
			<table width=500 border=0 cellpadding=2 cellspacing=0>
				<tr> 
					<td class='normal'><B>Date </B><%= FormatDateTime(Purchase_Date,2) %></td>
					<td class='normal'><B>Order ID </B><%= Store_Id %>-<%= Payment_ID %></td>
					<td class='normal'><B>Payment </B><%= Payment_Method %></td>

				</tr>
			</table>
			
			<hr>

			<table width=500 border=0 cellpadding=2 cellspacing=0>
				<tr>
					<td valign=top width=50% class='normal'>
						<B>Bill To</B><BR>
						<%= Site_Name %><BR>
						<% if Company <> "" then %>
						<%= Company %><BR>
						<% end if %>
                  <%= First_name %>&nbsp;<%= Last_Name %><BR>
                  <%= Address %><BR>
                  <%= City %>,&nbsp;<%= State %>&nbsp;<%= Zip %><BR>
                  <%= Country %><BR>
                  <%= Phone %><BR>
                  <%= Email %><BR>
					</TD>


				</tr>
			</table>
			
			<hr>
			
			<font CLASS='normal'>Order Details</font>









         <table cellspacing="0" cellpadding="0" border="1"  width="100%" >
	<tr>   

		<td class='normal'><STRONG>Quantity</STRONG></td>
		<td class='normal' width=30%><STRONG>Name</STRONG></td>
		<td class='normal'><STRONG>Price</STRONG></td>
		<td class='normal'><STRONG>Total</STRONG> </td>

	</tr>

		<tr>


				<td class='normal'>1</td>

    		<td class='normal'><%= Payment_Description %></td>
			<td class='normal'>$<%= FormatNumber(Grand_Total,2) %></td>
    		<td class='normal'>$<%= FormatNumber(Grand_Total,2) %></td>


		</tr>

</table>

<br>











			<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
			<tr>
				<TD width="201" height="1" colspan="4" class='normal'><STRONG>Total</STRONG> </TD>
				<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Grand_Total,2) %></STRONG></TD>
			</tr>
				<tr>
					<TD width="201" height="1" colspan="4" class='normal'><STRONG>Shipping and Handling</STRONG>  </TD>
				<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Shipping_Method_Price,2) %></STRONG></TD>
			</tr>
				<tr>
					<TD width="201" height="1" colspan="4" class='normal'><STRONG>Tax </STRONG></TD>
					<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Tax,2) %></STRONG></TD>
				</tr>
				<tr>
					<TD width="204" height="1" colspan="4" class='normal'><STRONG>Grand Total</STRONG> </TD>
					<TD width="71" height="1" class='normal'><STRONG>$<%= FormatNumber(Grand_Total,2) %></STRONG></TD>
			</tr>
			</TABLE>
		</td>
	</tr>


<% createFoot thisRedirect, 0%>

