<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

if Session("Super_User") <> 1 then
     Response.redirect "noaccess.asp"
end if

'check the data on the columns
'1- store_id
'2-Card Number
'3=date lessthan
'4=date greaterthan
'response.write request.form("lstCondition1" ) & "<br>"
select case cint(request.form("lstCondition1" ) )
	case 1
		cond1 = "Store_Id =" & request.form("txtValue1") 
	case 2
		cond1 = "Card_ending = '" & request.form("txtValue1")  &"'"
	case 3
		Cond1= "transDate < '" & request.form("txtValue1") & "'"
	case 4
		Cond1= "transDate > '" & request.form("txtValue1") & "'"
        case 5
                Cond1="email like '%"&request.form("txtValue1")&"%'"
        case 6
                Cond1="amount ="&request.form("txtValue1")
end select

'response.write cond1 & "<br>"& request.form("lstCondition2") &"<br>"
if request.form("lstAndOr") <> 0 then
	cond1 = cond1 & request.form("lstAndOr")
	select case cint(request.form("lstCondition2" ) )
		case 1
			cond1 = cond1 & "Store_Id =" & request.form("txtValue2") 
		case 2
			cond1 = cond1  & "Card_ending = '" & request.form("txtValue2")  &"'"
		case 3
			Cond1= cond1 & "transDate < '" & request.form("txtValue2") & "'"
		case 4
			Cond1= cond1 & "transDate > '" & request.form("txtValue2")  & "'"

	        case 5
                       Cond1=cond1&"email like '%"&request.form("txtValue2")&"%'" 
                case 6
                       Cond1=cond1&"amount ="&request.form("txtValue2")

        end select


end if


if request.form("txtValue1")   <> "" then
'			cond1 = request.form("Store_Id1")
			sql_select = "SELECT * from Sys_Payments where " &  cond1 &" order by transdate"



			set myfields=server.createobject("scripting.dictionary")

			Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)



			sql_select = "select Sys_Billing.Fee_Discount, Sys_Billing.Discount_Until, Sys_Billing.Amount, Sys_Billing.next_billing_date from Sys_Billing where sys_billing.store_id="&Store_Id1

			rs_store.open sql_select, conn_store, 1, 1
			if not rs_store.eof then
				Amount= rs_store("Amount")
				Fee_Discount = rs_store("Fee_Discount")
				Discount_Until = rs_store("Discount_Until")
				next_billing_date = rs_Store("next_billing_date")
			end if
			rs_store.close

			if Service_Type = 3 then
				Service = "Bronze"
			elseif Service_Type = 5 then
				Service = "Silver"
			elseif Service_Type = 7 then
				Service = "Gold"
			elseif Service_Type = 9 or Service_Type = 10 then
				Service = "Platinum"
			elseif Service_Type = 11 then
				Service = "Unlimited"

			else
				Service = "Unknown"
			end if

			if Discount_Until > Now() then
				Amount = Amount - Fee_Discount
			end if

			  if Discount_Until > BillDate then
				  Amount = Amount - Fee_Discount
			  end if
			  if Amount > 0 then
					sPaymentInfo = "Your next payment of <B>$"&Amount&"</b> for Easystorecreator "&Service&" Service will be automatically billed on <B>"&formatdatetime(next_billing_date,2)&"</b>"
			  end if

end if

sFormAction = "my_payments_admin.asp"
sSubmitName = "update_promotion_ids"
sTitle = "Store Payments"
sMenu = "account"
thisRedirect = "my_payments_admin.asp"
'sQuestion_Path = "import_export/my_account/prior_invoices.htm"
createHead thisRedirect %>

<tr bgcolor='#FFFFFF'>
	<td width="100%" colspan="3" align="center">
		<table>
		<tr bgcolor='#FFFFFF'>
		<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
		</tr>

			<tr bgcolor='#FFFFFF'>
				<td>
				<select name="lstCondition1">
					<option value=1>Store ID </option>
					<option value=2>Card # (last 4) </option>
					<option value=4 >Payment Date After </option>
					<option value=3>Payment Date Before </option>
					<option value=5>Email Like </option>
                                        <option value=6>Amount</option>

				</select>
				</td>
				<td><input type="text"  name = "txtValue1"></td>
				<td>
				<select name="lstAndOr">
					<option value=0>No 2nd Search</option>
					<option value=" and ">And</option>
					<option value=" or ">Or</option>
				</select>
				</td>
				
			</tr>
			<tr bgcolor='#FFFFFF'>
				<td>
				<select name="lstCondition2">
					<option value=2>Card # (last 4)</option>
					<option value=1>Store ID </option>
					<option value=3>Payment Date Before </option>
					<option value=4 >Payment Date After </option>
					<option value=5 >Email Like </option>
					<option value=6 >Amount </option>
				</select>
				</td>
				<td><input type="text"  name = "txtValue2"></td>
				<td>&nbsp;</td>
				
			</tr>

			<tr bgcolor='#FFFFFF'>
				
				<td colspan=2><input type="submit" value="Search"></td><td>&nbsp;</td>
				
			</tr>
					<tr bgcolor='#FFFFFF'>
		<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
		</tr>

		</table>
	</td>
</tr>


	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="13">

			
		<% 
			if noRecords = 1 then
		%>
		<table width="100%"  cellspacing="0" cellpadding=2>
			<tr bgcolor='#FFFFFF'>
			<td align="center">
			No Records were found
			</td>
			</tr>
		</table>

		<%
				elseif noRecords = 0 then
			      str_class=1
		%>
			<table width="100%" border="1" cellspacing="0" cellpadding=2>
			<tr bgcolor='#FFFFFF'>
			<td class=tablehead><b>Date</b> </td>
			<td class=tablehead><b>Store ID </b></td>
			<td class=tablehead><b>Card Ending </b></td>
			<td class=tablehead><b>Amount</b> </td>
			<td class=tablehead><b>Description</b> </td>
			<td class=tablehead><b>Details</b> </td>
			</tr>

		<%
					FOR rowcounter= 0 TO myfields("rowcount") 
						if str_class=1 then
							str_class=0
		                else
							str_class=1
		                end if 
			%>
			
				<tr bgcolor='#FFFFFF'>
					<td width="12%" class=<%= str_class %>><%= FormatDateTime(mydata(myfields("transdate"),rowcounter),2) %></td>
					<td width="12%" class=<%= str_class %>><%= mydata(myfields("store_id"),rowcounter) %></td>
					<td width="12%" class=<%= str_class %>><%= mydata(myfields("card_ending"),rowcounter) %></td>
					<td width="12%" class=<%= str_class %>>$<%=FormatNumber( mydata(myfields("amount"),rowcounter),2) %></td>
					<td width="40%" class=<%= str_class %>><%=mydata(myfields("payment_description"),rowcounter) %></td>
					<td width="12%" class=<%= str_class %>><a href=Payment_details.asp?store_id1=<%=mydata(myfields("store_id"),rowcounter)%>&Payment_Id=<%=mydata(myfields("payment_id"),rowcounter) %> class=link>View</a></td>
					</tr>

			<% Next %>
			<% end if %>
			</table>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'><td><BR><%= sPaymentInfo %></td></tr>


<% createFoot thisRedirect, 0%>


