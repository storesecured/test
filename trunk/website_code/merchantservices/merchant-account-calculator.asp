<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="recommended-merchant-account"

if request.form <> "" then
   monthly_sales = request.form("monthly_sales")
   if not isNumeric(monthly_sales) then
      monthly_sales = 1000
   end if
   average_sale = request.form("average_sale")
   if not isNumeric(average_sale) then
      average_sale = 100
   end if
   app_fee_spread = request.form("app_fee_spread")
   if not isNumeric(app_fee_spread) then
      app_fee_spread = 12
   end if
   app_fee = request.form("app_fee")
   if not isNumeric(app_fee) then
      app_fee = 0
   end if
   app_fee2 = request.form("app_fee2")
   if not isNumeric(app_fee2) then
      app_fee2 = 0
   end if
   monthly_fee = request.form("monthly_fee")
   if not isNumeric(monthly_fee) then
      monthly_fee = 0
   end if
   monthly_fee2 = request.form("monthly_fee2")
   if not isNumeric(monthly_fee2) then
      monthly_fee2 = 0
   end if
   annual_fee = request.form("annual_fee")
   if not isNumeric(annual_fee) then
      annual_fee = 0
   end if
   annual_fee2 = request.form("annual_fee2")
   if not isNumeric(annual_fee2) then
      annual_fee2 = 0
   end if
   discount_fee = request.form("discount_fee")
   if not isNumeric(discount_fee) then
      discount_fee = 0
   end if
   discount_fee2 = request.form("discount_fee2")
   if not isNumeric(discount_fee2) then
      discount_fee2 = 0
   end if
   trans_fee = request.form("trans_fee")
   if not isNumeric(trans_fee) then
      trans_fee = 0
   end if
   trans_fee2 = request.form("trans_fee2")
   if not isNumeric(trans_fee2) then
      trans_fee2 = 0
   end if
   min_fee = request.form("min_fee")
   if not isNumeric(min_fee) then
      min_fee = 0
   end if
   min_fee2 = request.form("min_fee2")
   if not isNumeric(min_fee2) then
      min_fee2 = 0
   end if
   
   num_transactions = monthly_sales / average_sale

   monthly_app1 = app_fee / app_fee_spread
   monthly_app2 = app_fee2 / app_fee_spread

   monthly_annual1 = annual_fee / 12
   monthly_annual2 = annual_fee2 / 12

   monthly_fees1 = (monthly_sales * discount_fee / 100) + (num_transactions * trans_fee)
   monthly_fees2 = (monthly_sales * discount_fee2 / 100) + (num_transactions * trans_fee2)
   
   if int(min_fee) > monthly_fees1 then
      monthly_fees1 = min_fee
   end if

   if int(min_fee2) > monthly_fees2 then
      monthly_fees2 = min_fee2
   end if
   
   total1 = monthly_fees1 + monthly_app1 + monthly_fee

   total2 = monthly_fees2 + monthly_app2 + monthly_fee2

   if total1 > total2 then
      lowest_total = total2
      higher_total = total1
      diff_total = total1 - total2
      lowest_answer = 2
   else
      lowest_total = total1
      higher_total = total2
      diff_total = total2 - total1
      lowest_answer = 1
   end if
   
   set myfields=server.createobject("scripting.dictionary")
	myfields.add "1_discount", 2.19
	myfields.add "1_trans",.28
	myfields.add "1_app",0
	myfields.add "1_annual",0
	myfields.add "1_monthly",25
	myfields.add "1_min",0
	myfields.add "1_url","https://cornerstonepaymentsystems.com/cpsagents/appwizard/autosid.asp?id=157318083"
  myfields.add "2_discount", 2.49
	myfields.add "2_trans",.35
	myfields.add "2_app",0
	myfields.add "2_annual",0
	myfields.add "2_monthly",20
	myfields.add "2_min",0
	myfields.add "2_url","https://cornerstonepaymentsystems.com/cpsagents/appwizard/autosid.asp?id=157586134"
	myfields.add "3_discount", 2.99
	myfields.add "3_trans",.50
	myfields.add "3_app",0
	myfields.add "3_annual",0
	myfields.add "3_monthly", 15
	myfields.add "3_min",0
	myfields.add "3_url","https://cornerstonepaymentsystems.com/cpsagents/appwizard/autosid.asp?id=156920134"
	myfields.add "4_discount", 5.5
	myfields.add "4_trans",.45
	myfields.add "4_app",49
	myfields.add "4_annual",0
	myfields.add "4_monthly",0
	myfields.add "4_min",0
	myfields.add "4_url","http://www.2checkout.com/cgi-bin/aff.2c?affid=74224"

	newlowest=9999999
	lowestplan=0
	for counter=1 to 4
	  monthly_fees = (monthly_sales * myfields(counter&"_discount") / 100) + (num_transactions * myfields(counter&"_trans"))

     if (myfields(counter&"_min")) > (monthly_fees) then
	     monthly_fees = (myfields(counter&"_min"))
	  end if
	  monthly_app = myfields(counter&"_app") / app_fee_spread
     monthly_annual = annual_fee / 12

	  total = monthly_fees + monthly_app + myfields(counter&"_monthly")

	  if lowest_total => total and total < newlowest then
         newlowest=total
         lowestplan=counter
     end if
   next

end if
%>
<!--#include file="header.asp"-->
<H1>Merchant Account Calculator</h1>
<form method=post>
<table>
<tr><td>Monthly Sales</td><td><input type=text name=monthly_sales size=5 value=<%= monthly_sales %>></td></tr>
<tr><td>Average Sale Price</td><td><input type=text name=average_sale size=5 value=<%= average_sale %>></td></tr>
<tr><td>Spread Application Fee</td><td><input type=text name=app_fee_spread size=5 value=<%= app_fee_spread %>>months</td></tr>
</table>
<BR>
<table>
<tr><td><B>Provider 1</b></td><td align=center>vs</td><td><b>Provider 2</b></td></tr>
<tr><td><input name=app_fee size=5 value=<%= app_fee %>></td><td align=center>Application or Setup Fee</td><td><input type=text name=app_fee2 size=5 value=<%= app_fee2 %>></td></tr>
<tr><td><input name=monthly_fee size=5 value=<%= monthly_fee %>></td><td align=center>Monthly, Gateway and/or Statement Fee</td><td><input type=text name=monthly_fee2 size=5 value=<%= monthly_fee2 %>></td></tr>
<tr><td><input name=annual_fee size=5 value=<%= annual_fee %>></td><td align=center>Annual Fee</td><td><input type=text name=annual_fee2 size=5 value=<%= annual_fee2 %>></td></tr>
<tr><td><input name=discount_fee size=5 value=<%= discount_fee %>>%</td><td align=center>Discount Rate</td><td><input type=text name=discount_fee2 size=5 value=<%=discount_fee2 %>>%</td></tr>
<tr><td><input name=trans_fee size=5 value=<%= trans_fee %>></td><td align=center>Transaction Fee</td><td><input type=text name=trans_fee2 size=5 value=<%= trans_fee2 %>></td></tr>
<tr><td><input name=min_fee size=5 value=<%= min_fee %>></td><td align=center>Monthly Minimum</td><td><input type=text name=min_fee2 size=5 value=<%= min_fee2 %>></td></tr>
<tr><td><B><%= formatnumber(total1,2) %></b></td><td align=center><B>Total</b></td><td><B><%= formatnumber(total2,2) %></b></td></tr>
<tr><td align=center colspan=3><input type=submit size=5 name=Compare></td></tr>
<tr><td align=center colspan=3><B>Best Deal Provider is #<%=lowest_answer %></b><BR>Lowest price $<%=formatnumber(lowest_total,2)  %></td></tr>
<% if newlowest<>9999999 and request.form <> "" then %>
<tr><td align=center colspan=3><BR><B>You could save money with the following plan</b><BR>Lower price $<%=formatnumber(newlowest,2)  %>
<% response.write "<BR>"&myfields(lowestplan&"_discount")&"% Discount Rate"
	response.write "<BR>$"&formatnumber(myfields(lowestplan&"_trans"),2)&" Transaction Fee"
	response.write "<BR>$"&formatnumber(myfields(lowestplan&"_app"),2)&" Application Fee"
	response.write "<BR>$"&formatnumber(myfields(lowestplan&"_annual"),2)&" Annual Fee"
	response.write "<BR>$"&formatnumber(myfields(lowestplan&"_monthly"),2)&" Monthly Fee"
	response.write "<BR>$"&formatnumber(myfields(lowestplan&"_min"),2)&" Monthly Minimum"
   response.write "<BR><a target=_blank href='"&myfields(lowestplan&"_url")&"'>Apply Now</a>"
   set myfields=nothing
%>
</td></tr>
<% end if %>
<tr><td colspan=3>Click here to view our list of <a href=recommended-merchant-accounts.asp>recommended merchant accounts</a>.</td></tr>

</table>





</form>
<!--#include file="footer.asp"-->



