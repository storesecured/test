<!--#include virtual="common/connection.asp"-->

<%

' IF CALLING FROM AFFILIATES' PAGES, USE AFFILIATES_SESSION.ASP
' FOR SESSION MANAGEMENT
RUN_FROM_AFFILIATES = TRUE

if Session("AFFILIATE_SESSION")="TRUE" then
	%><!--#include file="affiliates_session.asp"--><%
else
	'AFFILIATE IS NOT LOGED IN,
	'REDIRECT AFFILIATE TO LOGIN WINDOW
	Session("Is_Login") = False
	Response.Redirect "affiliates_login.asp"
end if

dim tot_affs
dim af_codes()
dim tot_clicks()
dim tot_ords()
dim tot_vals()
dim tot_payout()
dim use_ltotals

Affiliate_amount = Session("Affiliate_amount")
Affiliate_type = Session("Affiliate_type")

checked1=""
checked2=""
checked3=""

'REPORT TYPE
'	  1 = DAILY
'	  2 = WEEKLY (DEFAULT)
'	  3 = MONTHLY
select case request.form("REP_TYPE")
	case "1"
		checked1="checked"
	case "2"
		checked2="checked"
	case "3"
		checked3="checked"
	case else
		checked2="checked"
end select

'REPORT'S START / END DATES
end_date = dateserial(year(now()),month(now()),day(now()))
end_date = dateadd("d",1,end_date)
start_date = dateadd("ww",-2,end_date)
		
if request.form("start_date")<>"" and request.form("end_date")<>"" then
	if isdate(request.form("start_date")) and isdate(request.form("end_date")) then
		if cdate(request.form("start_date"))<cdate(request.form("end_date")) then
			start_date = cdate(request.form("start_date"))
			end_date = cdate(request.form("end_date"))
		end if
	end if
end if

'SELECTING THE AFFILIATE(S)
sel_affi = ""
sql_select_affiliates = "select * from store_affiliates where store_id=-1"
if request.form("Afiliates")<>"" then
	if (inStr(request.form("Afiliates"),"-1")<=0) then
		sel_affi = request.form("Afiliates")
		sql_select_affiliates = "select * from store_affiliates where Affiliate_ID in ("&sel_affi&") and store_id="&store_id
	else
		sel_affi = "-1"
		sql_select_affiliates = "select * from store_affiliates where store_id="&store_id
	end if
else
	if request.querystring("aff_id")<>"" then
		sel_affi = request.querystring("aff_id")
		sql_select_affiliates = "select * from store_affiliates where Affiliate_ID in ("&sel_affi&") and store_id="&store_id
	end if
end if


sel_affi = Session("AFFILIATE_AFFILIATE_ID")
sql_select_affiliates = "select * from store_affiliates where Affiliate_ID in ("&sel_affi&") and store_id="&store_id

sFormAction = "Affiliates_Reports.asp"
sTitle = "Affiliates Report"
thisRedirect = "affiliates_reports.asp"

%>


<script language="JavaScript">
function goDetails(affID)
{
	document.forms[0].action = 'affiliates_details.asp';
	document.forms[0].AFF_DET_CODE.value = affID;
	document.forms[0].submit();
}
</script>
 
<body>
<form method="POST">


	<!--#include file="affiliates_header.asp"-->


<tr><td><table>
<input type="hidden" name="AFF_DET_CODE" value="">
<input type="hidden" name="AFF_DET_SDATE" value="<%= start_date %>">
<input type="hidden" name="AFF_DET_EDATE" value="<%= end_date %>">

		
					<tr>
						<td width="1%" nowrap><font face="Arial" size="2">Start Date </font></td>
						<td width="99%"><font face="Arial" size="2"><input type="text" name="START_DATE" size="20" value="<%= start_date %>"></font></td>
					</tr>
		
					<tr>
						<td width="1%" nowrap><font face="Arial" size="2">End Date </font></td>
						<td width="99%"><font face="Arial" size="2"><input type="text" name="END_DATE" size="20" value="<%= end_date %>"></font></td>
					</tr>
		
					<tr>
						<td width="1%" nowrap><font face="Arial" size="2">Report Type</font></td>
						<td width="99%"><font face="Arial" size="2">
							<input type="radio" name="REP_TYPE" value="1" <%= checked1 %>>Daily <br>
							<input type="radio" name="REP_TYPE" value="2" <%= checked2 %>>Weekly <br>
							<input type="radio" name="REP_TYPE" value="3" <%= checked3 %>>Monthly </font></td>
					</tr>
		
					<tr>
						<td width="1%" nowrap><font face="Arial" size="2">&nbsp;</font></td>
						<td width="99%"><font face="Arial" size="2"><br>
							<input type="submit" name="Generate_Report" value="Get Report"><br>&nbsp;</font></td>
					</tr>
				</table>
				</td>
		</tr>

		<%' FOR EACH AFFILIATE, COMPUTE NUMBER OF CLICKS / ORDERS  %>
		<% rs_Store.open sql_select_affiliates,conn_store,1,1  %>
		<% if not rs_Store.eof then %>
		<tr>
				<td width="100%" colspan="4" height="13">
				<table border="1" width="100%" height="1" cellspacing="2" cellpadding="2">
					<tr>
						<td rowspan="2" align="center"><font face="Arial" size="1"><b>Period</b></font></td>
							<% tot_affs = 0 %>
							<% do while not rs_Store.eof %>
								<% rs_Store.movenext %>
								<% tot_affs = tot_affs + 1 %>
							<% loop %>
							<% rs_Store.movefirst %>
							
							<% if tot_affs>1 then %>
								<% use_ltotals = true %>
							<% else %>
								<% use_ltotals = false %>
							<% End If %>
							
							<% redim af_codes(tot_affs-1) %>
							<% redim tot_clicks(tot_affs-1) %>
							<% redim tot_ords(tot_affs-1) %>
							<% redim tot_vals(tot_affs-1) %>
							<% redim tot_payout(tot_affs-1) %>
							<% cpos = 0 %>
			
							<% do while not rs_Store.eof %>
								<td align="center" colspan="4"
									<% if cpos mod 2 = 1 then %>
										bgcolor="#DDDDD"
									<% End If %>
									><font face="Arial" size="1"><b><%= rs_store("Contact_Name") %><br>(<%= rs_store("Code") %>)</b></font></td>
									<% af_codes(cpos) = clng(rs_store("Affiliate_ID"))%>
									<% tot_clicks(cpos) = 0 %>
									<% tot_ords(cpos) = 0 %>
									<% tot_vals(cpos) = 0 %>
									<% tot_payout(cpos) = 0 %>
									<% cpos = cpos + 1 %>
									<% rs_Store.movenext %>
							<% loop %>
							<% rs_Store.close %>
							<% If use_ltotals then %>
								<td align="center" colspan="4" bgcolor="#AAAAA"><font face="Arial" size="1"><b>Totals</b></font></td>
							<% End If %>
					</tr>
					
					<tr>
						<% for i=1 to tot_affs %>
							<% If i mod 2 = 0 then %>
								<% bgcolor = "bgcolor=""#DDDDD""" %>
							<% Else %>
								<% bgcolor ="" %>
							<% End If %>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><b>Clicks</b></font></td>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><b>Orders</b></font></td>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><b>Order Totals</b></font></td>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><b>Payout</b></font></td>
						<% next %>
						<% If use_ltotals then %>
							<td align="center" bgcolor="#AAAAA"><font face="Arial" size="1"><b>Clicks</b></font></td>
							<td align="center" bgcolor="#AAAAA"><font face="Arial" size="1"><b>Orders</b></font></td>
							<td align="center" bgcolor="#AAAAA"><font face="Arial" size="1"><b>Order Totals</b></font></td>
							<td align="center" bgcolor="#AAAAA"><font face="Arial" size="1"><b>Payout</b></font></td>
						
		 <% End If %>
					</tr>
		
	
					<% cend_date = false %>
					<% will_exit = false %>
					<% cur_date = start_date %>
					<% prev_cur_date = start_date %>
					<% If checked1<>"" then %>
						<% cur_date = dateadd("d",1,cur_date) %>
					<% End If %>
					<% If checked2<>"" then %>
						<% cur_date = dateadd("ww",1,cur_date) %>
					<% End If %>
					<% If checked3<>"" then %>
						<% cur_date = dateadd("m",1,cur_date) %>
					<% End If %>
		
		
					<% crow=1 %>
					<% do while not cend_date %>
					<tr
						<% If crow mod 2 = 0 then %>
							bgcolor="#DDDDD"
						<% End If %>
						>
						<td align="center"><font face="Arial" size="1">
						<% If checked1<>"" then %>
							<%= month(prev_cur_date)&"/"&day(prev_cur_date)&"/"&year(prev_cur_date) %>
						<% else %>
							<%= month(prev_cur_date)&"/"&day(prev_cur_date)&"/"&year(prev_cur_date) %><br>
							<%= month(cur_date)&"/"&day(cur_date)&"/"&year(cur_date) %>
						<% End If %>
						</font></td>
				
						<% l_clicks = 0 
						 l_ords = 0
						 l_vals = 0
						 l_payout = 0
						 for i=1 to tot_affs 
							 If i mod 2 = 0 then 
								 bgcolor = "bgcolor=""#DDDDD""" 
							 Else  
								 bgcolor ="" 
							 End If 
							 sql_select_c1 = "select count(SS.store_id) as totClicks from store_shoppers SS, store_affiliates SA where SA.Affiliate_ID="&af_codes(i-1)&" and SS.CAME_FROM=SA.Code and SS.SYS_Created>=#"&prev_cur_date&"# and SS.SYS_Created<#"&cur_date&"# and SS.store_id="&store_id&" and SA.Store_ID="&store_id
							 if Store_Database="Ms_Sql" then 
								 sql_select_c1=replace(sql_select_c1,"#","'") 
							 end if   
							 rs_Store.open sql_select_c1,conn_store,1,1	
							 If rs_Store("totClicks")="" or isnull(rs_Store("totClicks")) then 
								 c_clicks = 0 
							 Else 
								 c_clicks = rs_Store("totClicks") 
							 End If 
							 rs_Store.close 
							 tot_clicks(i-1) = tot_clicks(i-1) + c_clicks 
							 l_clicks = l_clicks + c_clicks 

							 sql_select_c2 = "select count(SP.store_id) as totOrd, sum(SP.Total) as totAm from Store_Affiliates SA, Store_Purchases SP where SP.verified=1 and SA.Affiliate_ID="&af_codes(i-1)&" and SP.CAME_FROM=SA.Code and SP.SYS_Created>=#"&prev_cur_date&"# and SP.SYS_Created<#"&cur_date&"# and SA.store_id="&store_id&" and SP.store_id="&Store_id
							 if Store_Database="Ms_Sql" then
								sql_select_c2=replace(sql_select_c2,"#","'")
							 end if

							 rs_Store.open sql_select_c2,conn_store,1,1	
							 If rs_Store("totOrd")="" or isnull(rs_Store("totOrd")) then 
								 c_ords = 0 
							 Else  
								 c_ords = rs_Store("totOrd") 
							 End If 
							 If rs_Store("totAm")="" or isnull(rs_Store("totAm")="") then 
								 c_vals = 0 
							 Else 
								 c_vals = rs_Store("totAm") 
							 End If 
							 rs_Store.close 
							 tot_ords(i-1) = tot_ords(i-1) + c_ords 
							 tot_vals(i-1) = tot_vals(i-1) + c_vals 
							
							 l_ords = l_ords + c_ords 
							 l_vals = l_vals + c_vals 
							 if Affiliate_type = 1 then
								c_payout = c_vals * Affiliate_amount / 100
								l_payout = l_vals * Affiliate_amount / 100
							 else
                        c_payout = c_clicks * Affiliate_amount
								l_payout = l_clicks * Affiliate_amount
							 end if
							 tot_payout(i-1) = tot_payout(i-1) + c_payout %>

							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><%= c_clicks %></font></td>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><%= c_ords %></font></td>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(c_vals,2) %></font></td>
							<td align="center" <%= bgcolor %>><font face="Arial" size="1"><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(c_payout,2) %></font></td>
						<% next %>
			
						<% If use_ltotals then %>
							<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><%= l_clicks %></font></td>
							<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><%= l_ords %></font></td>
							<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(l_vals,2) %></font></td>
							<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(l_payout,2) %></font></td>
						<% End If %>
		
						<% prev_cur_date = cur_date %>
						<% If checked1<>"" then %>
							<% cur_date = dateadd("d",1,cur_date) %>
						<% End If %>
						<% If checked2<>"" then %>
							<% cur_date = dateadd("ww",1,cur_date) %>
						<% End If %>
						<% If checked3<>"" then %>
							<% cur_date = dateadd("m",1,cur_date) %>
						<% End If %>
		
						<% if cur_date>end_date then %>
							<% cend_date = true %>
						<% End If %>
						<% crow = crow + 1 %>
					</tr>
				<% loop %>
		
					<tr bgcolor="#AAAAAA">
						<td align="center"><font face="Arial" size="1"><b><br>Totals<br>&nbsp;</b></font></td>
							<% gt_clicks = 0 %>
							<% gt_ords = 0 %>
							<% gt_vals = 0 %>
							<% gt_payout = 0 %>
							<% for i=1 to tot_affs %>
								<% gt_clicks = gt_clicks + tot_clicks(i-1) %>
								<% gt_ords = gt_ords + tot_ords(i-1) %>
								<% gt_vals = gt_vals + tot_vals(i-1) %>
								<% gt_payout = gt_payout + tot_payout(i-1) %>
								<td align="center"><font face="Arial" size="1"><b><%= tot_clicks(i-1) %></b></font></td>
								<td align="center"><font face="Arial" size="1"><b><%= tot_ords(i-1) %></b></font></td>
								<td align="center"><font face="Arial" size="1"><b><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(tot_vals(i-1),2) %></b></font></td>
								<td align="center"><font face="Arial" size="1"><b><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(tot_payout(i-1),2) %></b></font></td>
							<% next %>
							<% If use_ltotals then %>
								<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><b><%= gt_clicks %></b></font></td>
								<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><b><%= gt_ords %></b></font></td>
								<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><b><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(gt_vals,2) %></b></font></td>
								<td align="center" bgcolor="#AAAAAA"><font face="Arial" size="1"><b><%= Session("AFFILIATE_CURRENCY") %><%= formatnumber(gt_payout,2) %></b></font></td>
							<% End If %>
					</tr>

					<tr bgcolor="#DDDDDD">
						<td align="center"><font face="Arial" size="1"><b><br>More Details<br>&nbsp;</b></font></td>
						<% for i=1 to tot_affs %>
							<td align="center" colspan="4"><font face="Arial" size="1"><b><a href="JavaScript:goDetails('<%= af_codes(i-1) %>');">Click and Order Details</a></b></font></td>
						<% next %>
						<% If use_ltotals then %>
							<td align="center" colspan="4"><font face="Arial" size="1"><b>&nbsp;</b></font></td>
						<% End If %>
					</tr>


	<% else %>
		<% rs_Store.close %>
	<% End If %>
</form>
