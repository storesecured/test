<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
calendar=1
'SET THE DEFAULT VALUES
Start_date =  formatdatetime(DateAdd("m", -1, Now()),2)
End_date = formatdatetime(Now(),2)

'COMPUTE THE REPORT FIELDS
if Request.Form("View_Date") <> "" then 
	Start_date =  Request.Form("Start_date")
	End_date = Request.Form("End_date")
	if Request.Form("Payment_Method") = "All" then
		Payment_Method_sql_str = ""
	else
		Payment_Method_sql_str = "Payment_Method = '"&Request.Form("Payment_Method")&"' AND "
	end if
	End_date1 = DateAdd("d", 1, End_date)
	if Request.Form("Show") = "All" then 
		sql_select_orders = "select distinct Verified, Store_Purchases.Purchase_Date, Store_Purchases.oid, Store_Purchases.cid, Tax, Shipping_Method_Price, Payment_Method, Store_Purchases.Total,WholeSale_Total,Coupon_Amount, [Total] - [WholeSale_Total] - [Coupon_Amount] as Profit_Loss From Store_Purchases Where "&Payment_Method_sql_str&" Store_Purchases.Purchase_Date < #"&End_date1&"# AND Store_Purchases.Purchase_Date > #"&Start_date&"# AND Store_Purchases.Store_id = "&Store_id&" ORDER BY Store_Purchases.OID DESC"
		sql_select_order_oid = "select Store_Purchases.oid From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where "&Payment_Method_sql_str&" Store_Purchases.Purchase_Date < #"&End_date1&"# AND Store_Purchases.Purchase_Date > #"&Start_date&"# AND Store_Purchases.Store_id = "&Store_id&" and Store_Transactions.Store_Id="&Store_Id
	else 
		Show = int(Request.Form("Show"))
		if Show = 2 then
		   sql_select_orders = "select distinct Verified, Store_Purchases.Purchase_Date, Store_Purchases.oid, Store_Purchases.cid, Tax, Shipping_Method_Price, Payment_Method, Store_Purchases.Total,WholeSale_Total,Coupon_Amount, [Total] - [WholeSale_Total] - [Coupon_Amount] as Profit_Loss From Store_Purchases where "&Payment_Method_sql_str&" returned<>0 AND Store_Purchases.Purchase_Date < #"&End_date1&"# AND Store_Purchases.Purchase_Date > #"&Start_date&"# AND Store_Purchases.Store_id = "&Store_id
    		sql_select_order_oid = "select distinct Store_Purchases.oid From Store_Purchases Where "&Payment_Method_sql_str&" Store_Purchases.Purchase_Date < #"&End_date1&"# AND Store_Purchases.Purchase_Date > #"&Start_date&"# AND Store_Purchases.Store_id = "&Store_id&" and Returned<>0 and Store_Transactions.Store_Id="&Store_Id
		else
    		sql_select_orders = "select distinct Verified, Store_Purchases.Purchase_Date, Store_Purchases.oid, Store_Purchases.cid, Tax, Shipping_Method_Price, Payment_Method, Store_Purchases.Total,WholeSale_Total,Coupon_Amount, [Total] - [WholeSale_Total] - [Coupon_Amount] as Profit_Loss From Store_Purchases Where "&Payment_Method_sql_str&" Verified = "&Show&" and returned is Null AND Store_Purchases.Purchase_Date < #"&End_date1&"# AND Store_Purchases.Purchase_Date > #"&Start_date&"# AND Store_Purchases.Store_id = "&Store_id
    		sql_select_order_oid = "select distinct Store_Purchases.oid From Store_Purchases Where "&Payment_Method_sql_str&" Verified = "&Show&" AND Store_Purchases.Purchase_Date < #"&End_date1&"# AND Store_Purchases.Purchase_Date > #"&Start_date&"# AND Store_Purchases.Store_id = "&Store_id&" and Returned is Null"
	   end if
   end if

elseif Request.Form("View_Oid") <> "" then
	if Request.Form("Oid") = "" then
		response.redirect "admin_error.asp?message_id=21"
	elseif not isnumeric(Request.Form("Oid")) then
		response.redirect "admin_error.asp?message_id=22"
	else
		Oid = int(Request.Form("Oid"))
	end if
	sql_select_orders = "select distinct Verified, Store_Purchases.Purchase_Date, Store_Purchases.oid, Store_Purchases.cid, Tax, Shipping_Method_Price, Payment_Method, Store_Purchases.Total,WholeSale_Total,Coupon_Amount, [Total] - [WholeSale_Total] - [Coupon_Amount] as Profit_Loss From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.Oid = "&Oid&" AND Store_Purchases.Store_id = "&Store_id&" and Store_Transactions.Store_Id="&Store_Id
	sql_select_order_oid = "select distinct Store_Purchases.oid From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.Oid = "&Oid&" AND Store_Purchases.Store_id = "&Store_id&" and Store_Transactions.Store_Id="&Store_Id

elseif Request.Form("View_Cid") <> "" then 
	if Request.Form("Cid") = "" then
		response.redirect "admin_error.asp?message_id=21"
	elseif not isnumeric(Request.Form("Cid")) then
		response.redirect "admin_error.asp?message_id=22"
	else
		Cid = int(Request.Form("Cid"))
	end if
	sql_select_orders = "select distinct Verified, Store_Purchases.Purchase_Date, Store_Purchases.oid, Store_Purchases.cid, Tax, Shipping_Method_Price, Payment_Method, Store_Purchases.Total,WholeSale_Total,Coupon_Amount, [Total] - [WholeSale_Total] - [Coupon_Amount] as Profit_Loss From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.CCid = "&Cid&" AND Store_Purchases.Store_id = "&Store_id&" and Store_Transactions.Store_Id="&Store_Id
	sql_select_order_oid = ""
	sql_select_order_oid = "select distinct Store_Purchases.oid From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.CCid = "&Cid&" AND Store_Purchases.Store_id = "&Store_id&" and Store_Transactions.Store_Id="&Store_Id

end if

sql_select_orders_totals = "select sum(store_purchases.tax) as sum_tax, sum(store_purchases.Shipping_Method_Price) as sum_Shipping_Method_Price, sum(store_purchases.Total) as sum_Total, sum(store_purchases.WholeSale_Total) as sum_WholeSale_Total, sum(store_purchases.Coupon_Amount) as Sum_Coupon_Amount, sum([Total] - [WholeSale_Total] - [Coupon_Amount]) as Profit_Loss From Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.Purchase_Date < '"&Start_date&"' AND Store_Purchases.Purchase_Date > '"&End_date1&"' AND Store_Purchases.Store_id = "&Store_ID&" and Store_Transactions.Store_Id="&Store_Id

TotalTax = 0
TotalShipping = 0
TotalTotal = 0
TotalWholesale = 0
TotalCoupon = 0
TotalProfitLoss = 0
TotalOrders = 0

sFormAction = "reports_Order_Break_down.asp"
sName = "reports_Order_Break_down"
sTitle = "Orders Alternate View"
sFullTitle = "<a href=orders.asp class=white>Orders</a> > Alternate View"
thisRedirect = "reports_Order_Break_down.asp"
sMenu = "orders"
sQuestion_Path = "reports/order_break_down.htm"
createHead thisRedirect
%>


	 
	<TR bgcolor='#FFFFFF'>
	<td class="inputname">Order ID</td>
			<td class="inputvalue">
			<input type="text" name="Oid" size="13" value="<%= Request.Form("OID") %>">
			<% small_help "Order ID" %></td>
	</tr>
	<TR bgcolor='#FFFFFF'><td colspan=3 align=center><input name="View_Oid" type="submit" class="Buttons" value="Search Orders"></td></tr>
	<TR bgcolor='#FFFFFF'>
		<td class="inputname">Customer ID</td>
			<td class="inputvalue">
			<input type="text" name="Cid" size="13" value="<%= Request.Form("Cid") %>">
			<% small_help "Customer ID" %></td>
	 </tr>
	 <TR bgcolor='#FFFFFF'><td colspan=3 align=center><input name="View_Cid" type="button" class="Buttons" value="Search Orders" onclick="javascript:document.reports_Order_Break_down.submit()"></td></tr>
	<TR bgcolor='#FFFFFF'>

			<td class="inputname">Status</td>
			<td class="inputvalue">
			<select name="Show" size="1">
				<option 
					<% if request.form("Show")="" or request.form("Show")="All" then %>
						selected
					<% End If %>
					value="All">All</option>
				<option value="1"
					<% if request.form("Show")="1" then %>
						selected
					<% End If %>
					>Verified
				<option value="0"
					<% if request.form("Show")="0" then %>
						selected
					<% End If %>
					>Non Verified</option>
				<option value="2"
					<% if request.form("Show")="2" then %>
						selected
					<% End If %>
					>Returned</option>
			</select>
			<% small_help "Status" %></td></tr>
			<TR bgcolor='#FFFFFF'><td class="inputname">Payment Method</td>
			<td class="inputvalue">
			<% Call Disp_Payment %>
			<% small_help "Payment Method" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'><td class="inputname">Between</td>
			<td class="inputvalue">
      <SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
			<input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
      <A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			<input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
			<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
      <% small_help "Between" %></td></tr>

			<TR bgcolor='#FFFFFF'><td colspan=3 align=center><input name="View_Date" type="submit" class="Buttons" value="View"></td>
	</tr>
	<TR bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="74">

			
			<table width="100%" border="1" cellspacing="0" cellpadding=2 class="list">
				<%  if sql_select_orders <> "" then %>
					<% if Store_Database="Ms_Sql" then %>
						<% sql_select_orders = Replace(sql_select_orders, "#", "'") %>
					<% end if %>

					<% set myfields=server.createobject("scripting.dictionary")
						Call DataGetrows(conn_store,sql_select_orders,mydata,myfields,noRecords)
						Number_Of_Orders = myfields("rowcount")
					%>

					<TR bgcolor='#FFFFFF'>
						<td class=tablehead><b>Verified</b></td>
						<td class=tablehead><b>Purchase Date</b></td>
						<td class=tablehead><b>Order</b></td>
						<td class=tablehead><b>Payment Method</b></td>
						<td class=tablehead><b>Tax</b></td>
						<td class=tablehead><b>Shipping</b></td>
						<td class=tablehead><b>Revenue</b></td>
						<td class=tablehead><b>Cost</b></td>
						<td class=tablehead><font color="#FF0000"><b>Coupon Amount</b></font></td>
						<td class=tablehead><b>Profit/Loss</b></td>
					</tr>
					<% str_class=1
						if noRecords = 0 then
						FOR rowcounter= 0 TO myfields("rowcount")

					  if str_class=1 then
							str_class=0
					  else
							str_class=1
					  end if %>
						<TR bgcolor='#FFFFFF'>
						<% if mydata(myfields("verified"),rowcounter) = -1 then
							Verified = "Yes"
						Else
							Verified = "No"
						end if %>

							<td width="53" height="10" class="<%= str_class %>"><%= Verified %></td>
							<td width="53" height="10" class="<%= str_class %>"><%= mydata(myfields("purchase_date"),rowcounter) %></td>
							<td width="53" height="10" class="<%= str_class %>"><a class="link" href='Order_Details.asp?id=<%=mydata(myfields("oid"),rowcounter)%>'><%=mydata(myfields("oid"),rowcounter)%></a></td>
							<td width="54" height="10" class="<%= str_class %>"><%= mydata(myfields("payment_method"),rowcounter) %></td>
							<td width="54" height="10" class="<%= str_class %>"><%= Currency_Format_Function(mydata(myfields("tax"),rowcounter)) %></td>
							<td width="54" height="10" class="<%= str_class %>"><%= Currency_Format_Function(mydata(myfields("shipping_method_price"),rowcounter)) %></td>
							<td width="54" height="10" class="<%= str_class %>"><%= Currency_Format_Function(mydata(myfields("total"),rowcounter)) %></td>
							<td width="54" height="10" class="<%= str_class %>"><%= Currency_Format_Function(mydata(myfields("wholesale_total"),rowcounter)) %></td>
							<td width="54" height="10" class="<%= str_class %>"><%= Currency_Format_Function(mydata(myfields("coupon_amount"),rowcounter)) %></td>
							<td width="54" height="10" class="<%= str_class %>"><%= Currency_Format_Function(mydata(myfields("profit_loss"),rowcounter)) %></td>
							<% TotalTax = TotalTax + mydata(myfields("tax"),rowcounter)
								TotalShipping = TotalShipping + mydata(myfields("shipping_method_price"),rowcounter)
								TotalTotal = TotalTotal + mydata(myfields("total"),rowcounter)
								TotalWholesale = TotalWholesale + mydata(myfields("wholesale_total"),rowcounter)
								TotalCoupon = TotalCoupon + mydata(myfields("coupon_amount"),rowcounter)
								TotalProfitLoss = TotalProfitLoss + mydata(myfields("profit_loss"),rowcounter)
								TotalOrders = TotalOrders + 1
							%>
						</tr>
					<% Next %>
					<% End If %>

						<TR bgcolor='#FFFFFF'>
							<td width="54" height="14" bgcolor="#C0C0C0">&nbsp;</td>
							<td width="53" height="14" bgcolor="#C0C0C0">Totals</td>
							<td width="53" height="14" bgcolor="#C0C0C0"># orders</td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= TotalOrders %></td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= Currency_Format_Function(TotalTax) %></td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= Currency_Format_Function(TotalShipping) %></td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= Currency_Format_Function(TotalTotal) %></td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= Currency_Format_Function(TotalWholesale) %></td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= Currency_Format_Function(TotalCoupon) %></td>
							<td width="54" height="14" bgcolor="#C0C0C0"><%= Currency_Format_Function(TotalProfitLoss) %></td>
						</tr>
				<% end if %>
			</table>
		</td>
	</tr>

<% createFoot thisRedirect, 0%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");

</script>
