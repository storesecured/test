<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

calendar=1
'SET THE DEFAULT VALUES
Start_date =  formatdatetime(DateAdd("m", -1, Now()),2)
End_date = formatdatetime(Now() ,2)
if Request.Form("Start_date") <> "" then 
	Start_date =  Request.Form("Start_date")
	End_date = Request.Form("End_date")
end if 

End_date1 = DateAdd("d", 1, End_date)

sql_select = "select * from store_statistics where store_id="&Store_Id
rs_store.open sql_select,conn_store,1,1
   skip_hosts = rs_store("skip_hosts")
rs_store.close
skip_hosts=replace(skip_hosts," ",",")
if skip_hosts="" then
   skip_hosts="1"
end if

sql_cart = "SELECT Count(Store_ShoppingCart.Shopping_cart_id) AS CountOfShopping_cart_id FROM Store_ShoppingCart Where Sys_Created < '"&End_date1&"' AND Sys_Created > '"&Start_date&"' And Store_id = "&Store_id&" "


rs_Store.open sql_cart,conn_store,1,1
	rs_Store.MoveFirst	  
	CountOfShopping_cart_id = Rs_store("CountOfShopping_cart_id")
rs_store.close

sql_customers = "SELECT Count(Store_Customers.Cid) AS CountOfCid FROM Store_Customers WHERE Store_Customers.Record_type=0 AND Sys_Created < '"&End_date1&"' AND Sys_Created > '"&Start_date&"' And Store_id = "&Store_id&""


rs_Store.open sql_customers,conn_store,1,1
	rs_Store.MoveFirst	  
	CountOfCid = Rs_store("CountOfCid")
rs_store.close

sql_Purchases_totals = "select Count(Store_Purchases.Oid) as Count_Of_Oid, sum(tax) as sum_tax, sum(Shipping_Method_Price) as sum_Shipping_Method_Price, sum(Store_Purchases.Total) as sum_Total, sum(WholeSale_Total) as sum_WholeSale_Total, sum(Coupon_Amount) as Sum_Coupon_Amount, sum([Total] - [WholeSale_Total] - [Coupon_Amount]) as Profit_Loss From Store_Purchases Where Store_Purchases.Purchase_Date < '"&End_date1&"' AND Store_Purchases.Purchase_Date > '"&Start_date&"' And Store_Purchases.Store_id = "&Store_id&" AND Store_Purchases.Verified = 1 and Returned is Null"


rs_Store.open sql_Purchases_totals,conn_store,1,1
	Count_Of_Oid = Rs_store("Count_Of_Oid")
	sum_tax= Rs_store("sum_tax")
	sum_Shipping_Method_Price = Rs_store("sum_Shipping_Method_Price")
	sum_Total = Rs_store("sum_Total")
	sum_WholeSale_Total = Rs_store("sum_WholeSale_Total")
	Sum_Coupon_Amount = Rs_store("Sum_Coupon_Amount")
	Profit_Loss = Rs_store("Profit_Loss")
rs_store.close

sql_purchases_verified = "SELECT DISTINCT Store_Purchases.OID FROM Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.Purchase_Date < '"&End_date1&"' AND Store_Purchases.Purchase_Date > '"&Start_date&"' And Store_Purchases.Store_id = "&Store_id&" and Store_Purchases.Verified = 1 and Returned is Null;"


rs_Store.open sql_purchases_verified,conn_store,1,1
Number_Of_Verified_Orders = get_number_of_recordset(rs_Store)
rs_store.close
	sql_purchases_verified = "SELECT DISTINCT Store_Purchases.OID FROM Store_Purchases INNER JOIN Store_Transactions ON Store_Purchases.OID = Store_Transactions.OID Where Store_Purchases.Purchase_Date < '"&End_date1&"' AND Store_Purchases.Purchase_Date > '"&Start_date&"' And Store_Purchases.Store_id = "&Store_id&" and Store_Purchases.Verified = 0 and returned is Null;"
	


set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_purchases_verified,mydata,myfields,noRecords)
Number_Of_Non_Verified_Orders = myfields("rowcount")


'COMPUTE SOME STATISTICS
if Count_Of_Oid = 0 then
	Count_Of_Oid = 0
	sum_tax= 0
	sum_Shipping_Method_Price = 0
	sum_Total = 0
	sum_WholeSale_Total = 0
	Sum_Coupon_Amount = 0
	Profit_Loss = 0
	Average_Sale = 0
	Average_Oid_Per_Day = 0
	Expected_Orders_Per_Month = 0
	Expected_Orders_Per_Year = 0
	Expected_Sales_Per_Month = 0
	Expected_Sales_Per_Year = 0
else
	Average_Sale = Currency_Format_Function(sum_Total/Count_Of_Oid)
	Average_Oid_Per_Day = FormatNumber(Count_Of_Oid/DateDiff("d",Start_date,End_date1),2)
	Expected_Orders_Per_Month = FormatNumber(Average_Oid_Per_Day*30,2)
	Expected_Orders_Per_Year = FormatNumber(Average_Oid_Per_Day*365,2)
	Expected_Sales_Per_Month = Currency_Format_Function((sum_Total/Count_Of_Oid)*Expected_Orders_Per_Month)
	Expected_Sales_Per_Year = Currency_Format_Function((sum_Total/Count_Of_Oid)*Expected_Orders_Per_Year)
end if

if CountOfShopping_cart_id = 0 then
	Conversion_Ratio = 0
else
	Conversion_Ratio = FormatNumber(Count_Of_Oid/(CountOfShopping_cart_id+Count_Of_Oid)*100,2)
end if

sFormAction = "Reports_Totals1.asp"
sName = "Reports_Totals"
sTitle = "Orders Report"
sFullTitle = "<a href=orders.asp class=white>Orders</a> > Report"
thisRedirect = "Reports_Totals1.asp"
sMenu = "orders"
sSubmitName="View_Date"
sQuestion_Path = "reports/traffic_and_totals.htm"
createHead thisRedirect

%>



	<tr bgcolor='#FFFFFF'>
		<td class="inputname">Between</td>
			<td class="inputvalue" colspan=2>
			<SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
			<input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
      <A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			and <input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
			<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
      <% small_help "Between" %></td>
	</tr>
	<tr bgcolor='#FFFFFF'><td colspan=4 align=center><input name="<%= sSubmitName %>" type="submit" class="Buttons" value="Search Orders"></td></tr>
	
	<TR bgcolor='#D4DEE5'><td width='100%' colspan='4' height='13'><table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				
				<tr bgcolor='#FFFFFF'>
					<td width="226" class=0>Verified Sales SubTotal&nbsp;</td>
				<td width="428" class=0><%= Currency_Format_Function(sum_Total) %></td>
				<td width="428" class=0>Total paid for items, does not include shipping, taxes, coupons etc</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="226" class=1>Cost Total</td>
					<td width="428" class=1><%= Currency_Format_Function(sum_WholeSale_Total) %></td>
					<td width="428" class=1>Cost Total as reported based on cost per item</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="226" class=0>Coupon Total</td>
					<td width="428" class=0><%= Currency_Format_Function(Sum_Coupon_Amount) %></td>
					<td width="428" class=0>Total coupon discounts</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="226" class=1>Profit/Loss Total<br></td>
					<td width="428" class=1><%= Currency_Format_Function(Profit_Loss) %></td>
					<td width="428" class=1>Order total minus coupons and item cost</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="226" class=0>Shipping Total</td>
					<td width="428" class=0><%= Currency_Format_Function(sum_Shipping_Method_Price) %></td>
					<td width="428" class=0>Shipping Charges</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="226" class=1>Tax Total</td>
					<td width="428" class=1><%= Currency_Format_Function(sum_tax) %></td>
					<td width="428" class=1>Tax charges</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="299" class=0>&nbsp;</td>
					<td width="355" class=0>&nbsp;</td>
					<td width="428" class=0>&nbsp;</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="299" class=1># of abandoned carts&nbsp; </td>
					<td width="355" class=1><%= CountOfShopping_cart_id %></td>
					<td width="428" class=1>Put an item in the cart but did not complete purchase (last 30 days only)</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="299" class=0># of orders&nbsp;&nbsp; </td>
					<td width="355" class=0><%= Count_Of_Oid %></td>
					<td width="428" class=0>Actual completed orders</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="299" class=1># new registrations&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
					<td width="355" class=1><%= CountOfCid %></td>
					<td width="428" class=1>Users who filled out their name and address</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="299" class=0>&nbsp;</td>
					<td width="355" class=0>&nbsp;</td>
					<td width="428" class=0>&nbsp;</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="299" class=1>Conversion ratio&nbsp;</td>
					<td width="355" class=1><%= Conversion_Ratio %>%</td>
					<td width="428" class=1># of orders&nbsp;/ carts&nbsp;&nbsp;</td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="299" class=0>Average Sale&nbsp;</td>
					<td width="355" class=0><%= Average_Sale %></td>
					<td width="428" class=0>orders amount&nbsp;/ orders</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="299" class=1>Average orders per day</td>
					<td width="355" class=1><%= Average_Oid_Per_Day %></td>
					<td width="428" class=1>orders/days </td>
				</tr>
						
				<tr bgcolor='#FFFFFF'>
					<td width="299" class=0>Expected orders per month.</td>
					<td width="355" class=0><%= Expected_Orders_Per_Month %></td>
					<td width="428" class=0>average order total per 30 days</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="299" class=1>Expected orders per year</td>
					<td width="355" class=1><%= Expected_Orders_Per_Year %></td>
					<td width="428" class=1>average orders per 365 days</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="299" class=0>Expected amount of sales per month</td>
					<td width="355" class=0><%= Expected_Sales_Per_Month %></td>
					<td width="428" class=0>Average order amount</td>
				</tr>
					
				<tr bgcolor='#FFFFFF'>
					<td width="299" class=1>Expected amount of sales per year</td>
					<td width="355" class=1><%= Expected_Sales_Per_Year %></td>
					<td width="428" class=1>Average order amount</td>
				</tr>

			</table>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'><td colspan=4>For more detailed statistics on referrals, search engines, number of visitors etc please visit the statistics page from Reports-->Statistics</td></tr>
<% createFoot thisRedirect, 2%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");

</script>
