<!--#include file="include/header.asp"-->

<%

'RETRIEVE AND SET THE DATA RANGE
if Request.Form("Form_Name") = "" then
	End_date = FormatDateTime(now(),2)
	Start_Date = dateadd("yyyy",-1,End_Date)

else
	if NOT IsDate(Request.Form("Start_date")) OR NOT IsDate(Request.Form("End_date") &" 11:59pm") then
		fn_redirect Site_Name&"error.asp?message_id=28"
	end if
	Start_date =  Request.Form("Start_date")
	End_date = Request.Form("End_date")
end if
End_Date1 = Dateadd("d",1,End_Date)


'RETRIEVE THE LIST OF PAST ORDERS FOR THE CUSTOMER
filter_string = ""
if Request.Form("Verified_Filter")<>"2" and Request.Form("Verified_Filter")<>"" then
	Verified_Filter = Request.Form("Verified_Filter")
else
    Verified_Filter = -1
end if

if Request.Form("Shipped_Filter")<>"all" and Request.Form("Shipped_Filter")<>"" then
	Shipped_Filter=Request.Form("Shipped_Filter")
else
    Shipped_Filter="all"
end if

sql_select_orders = "wsp_purchases_past "&Store_Id&","&cid&",'"&Start_Date&"','"&End_Date1&"',"&Verified_Filter&",'"&Shipped_Filter&"';"
fn_print_debug sql_select_orders
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select_orders,mydata,myfields,noRecords)

%>


<table border="0" width="100%">
	<tr> 
		<td width="100%" class='normal'><b>Past Orders for <%= first_name &"&nbsp;"& last_name %>  </b></td>
	</tr>
</table>

<form action="Past_orders.asp" method="post" name="past_orders">
<input type="Hidden" name="Form_Name" value="past_orders">

<table border="1" cellPadding="5" cellSpacing="0" width="100%">
	<tr>
		<td class='normal'>
			<table cellspacing=0 cellpadding=2><tr><td class='normal'>Orders from
			<input name="Start_date" size="10" value="<%= FormatDateTime(Start_date,2) %>" maxlength=10 onKeyPress="return goodchars(event,'0123456789/')">
			to
			<input name="End_date" size="10" value="<%= FormatDateTime(End_date,2) %>" maxlength=10 onKeyPress="return goodchars(event,'0123456789/')">
			&nbsp;</td><td class='normal'><%= fn_create_action_button ("Button_image_View", "View", "View") %>
			</td></tr></table>Filter by orders verified
			<select name="Verified_Filter">
				<option 
					<% if Verified_Filter = "2" then %>
						selected
					<% end if %>
					value="2">All</option>
				<option
					<% if Verified_Filter = "1" then %>
						selected
					<% end if %>
					value="1">Yes</option>
				<option
					<% if Verified_Filter = "0" then %>
						selected
					<% end if %>
					value="0">No</option></select>
			<% if Show_shipping then %>
			<br>Filter by orders shipped
			<select name="Shipped_Filter">
				<option 
					<% if Shipped_Filter= "all" then %>
						selected
					<% end if %>
					value="all">All</option>
				<option 
					<% if Shipped_Filter = "yes" then %>
						selected
					<% end if %>
					value="yes">Yes</option>
				<option 
					<% if Shipped_Filter = "no" then %>
						selected
					<% end if %>
					value="no">No</option>
				<option 
					<% if Request.Form("Shipped_Filter") = "partial" then %>
						selected
					<% end if %>
					value="partial">Partial</option></select>
		<% end if %>
		</td>
	</tr>
	
<tr>
<td height="150" vAlign="top">
<table border="0" cellPadding="1" cellSpacing="2">
<tr> 
<td class='normal'><strong>Verified</strong></td>
<td class='normal'><strong>Date</strong></td>
<td class='normal'><strong>Order</strong></td>
<% if Show_shipping then %>
    <td class='normal'><strong>Shipped</strong></td>
<% end if %>
<td class='normal'><strong>Payment</strong></td>
<td class='normal'><strong>Total</strong></td>
<td class='normal'><strong>Status</strong></td>
<% if Enable_RMA then %>
    <td class='normal'><strong>&nbsp;</strong></td>
<% end if %>
<td class='normal'><strong>&nbsp;</strong></td>
</tr>
<% 
if noRecords = 0 then
    FOR rowcounter= 0 TO myfields("rowcount")
        Verified=mydata(myfields("verified"),rowcounter)
        Return_Date=mydata(myfields("return_date"),rowcounter)
        Return_Rma=mydata(myfields("return_rma"),rowcounter)
	    shippedpr=mydata(myfields("shippedpr"),rowcounter)
	    sOid=mydata(myfields("oid"),rowcounter)
	    payment_method=mydata(myfields("payment_method"),rowcounter)
	    grand_total=mydata(myfields("grand_total"),rowcounter)
	    purchase_date=mydata(myfields("purchase_date"),rowcounter)
	    if cint(Verified)<>cint(0) then
		    Verified = "Yes"
	    Else
		    Verified = "No"
	    end if

        if Return_Date <> "" and not isnull(Return_Date) then
	        Status = "Returned"
        elseif Return_Rma <> "" and Return_Rma<>Null then
	        Status = "RMA Req"
        elseif shippedpr = "yes" then
	        Status = "Shipped"
        elseif Verified = -1 then
		    Status = "Processing"
	    else
		    Status = "Ordered"
	    end if 
	    %>
        <tr>
        <td class='normal'><%= Verified %></td>
        <td class='normal'><%= FormatDateTime(Purchase_Date,2) %></td>
        <td><a href="Past_order_detail.asp?Soid=<%= sOid %>" class='link'><%= sOid %></td>
        <% if Show_shipping then %>
            <td class='normal'><%= ShippedPr %></td>
        <% end if %>
        <td class='normal'><%= payment_method %></td>
        <td class='normal' align=right><%= Currency_Format_Function(Grand_Total) %></td>
        <td class='normal'><%= Status %></td>
        <% if Enable_RMA and (Return_Rma = "" or isNull(Return_Rma)) then %>
            <td class='normal'><a href=rma.asp?oid=<%= sOid %> class=link>Return</a></td>
        <% end if %>
        </tr>
    <% next %>
<% end if %> 
</table>
</td>
</tr>
<input name="NoDelete" type="hidden" value="1"></FORM>
</table>
<%
sFormName = "past_orders"
sSubmitName = "View"
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("Start_date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_date","date","Please enter a valid end date.");
</script>

<!--#include file="include/footer.asp"-->
