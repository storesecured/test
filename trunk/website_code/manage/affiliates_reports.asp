<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
calendar=1
Server.ScriptTimeout = 150
' IF CALLING FROM AFFILIATES' PAGES, USE AFFILIATES_SESSION.ASP
' FOR SESSION MANAGEMENT
RUN_FROM_AFFILIATES = FALSE
Session("AFFILIATE_SESSION")="FALSE"
%><!--#include file="include\sub.asp"--><%


dim tot_affs
dim af_codes()
dim tot_clicks()
dim tot_ords()
dim tot_vals()
dim tot_payout()
dim use_ltotals

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
sql_select_affiliates = "select * from store_affiliates WITH (NOLOCK) where store_id=-1"
if request.form("Afiliates")<>"" then
	if (inStr(request.form("Afiliates"),"-1")<=0) then
		sel_affi = request.form("Afiliates")
		sql_select_affiliates = "select * from store_affiliates WITH (NOLOCK) where Affiliate_ID in ("&sel_affi&") and store_id="&store_id
	else
		sel_affi = "-1"
		sql_select_affiliates = "select * from store_affiliates WITH (NOLOCK) where store_id="&store_id
	end if
else
	if request.querystring("Id")<>"" or request.querystring("aff_id") <> "" then
		sel_affi = request.querystring("Id")
		if sel_affi = "" then
		sel_affi = request.querystring("aff_id")
		end if
		if not isNumeric(sel_affi) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		sql_select_affiliates = "select * from store_affiliates WITH (NOLOCK) where Affiliate_ID in ("&sel_affi&") and store_id="&store_id
	end if
end if

on error resume next

sFormAction = "Affiliates_Reports.asp"
sTitle = "Affiliates Report"
sFullTitle = "Marketing > <a href=affiliates_manager.asp class=white>Affiliates</a> > Report"
thisRedirect = "affiliates_reports.asp"
sMenu = "marketing"
sQuestion_Path = "reports/affiliates_1.htm"
createHead thisRedirect
if Service_Type < 7	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
  <% createFoot thisRedirect, 0%>
<% else %>




<script language="JavaScript">
function goDetails(affID)
{
	document.forms[0].action = 'affiliates_details.asp';
	document.forms[0].AFF_DET_CODE.value = affID;
	document.forms[0].submit();
}
</script>
 


<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='13'><table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
<input type="hidden" name="AFF_DET_CODE" value="">
<input type="hidden" name="AFF_DET_SDATE" value="<%= start_date %>">
<input type="hidden" name="AFF_DET_EDATE" value="<%= end_date %>">

						<tr bgcolor='#FFFFFF'>
							<td class="inputname">Affiliates</td>
							<td class="inputvalue" width="70%">
								<select name="Afiliates" multiple>
									<option value="-1"
									<% If cstr(sel_affi)="-1" or instr(", "&sel_affi&",","-1")>0 then %>
										selected
									<% End If %>
									>All</option>
								<% sql_all_aff = "select * from store_affiliates WITH (NOLOCK) where store_id="&store_id&" order by code" %>
								<% rs_store.open sql_all_aff,conn_store,1,1%>
								<% do while not rs_store.eof %>
									<option value="<%= rs_store("Affiliate_ID") %>"
									<% if instr(", "&sel_affi&",",cstr(rs_store("Affiliate_ID")))>0 or instr(", "&sel_affi&",","-1")>0 then %>
										selected
									<% End If %>
									><%= rs_store("Contact_Name") %> (<%= rs_store("Code") %>)</option>
									<% rs_store.movenext  %>
								<% loop %>
								<% rs_store.close %>
							</select><% small_help "Affiliates" %></td>
						</tr>

		
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Between</td>
						<td class="inputvalue" width="70%"><SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
			<input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
      <A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			<input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
			<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>

						<% small_help "Between" %></td>
					</tr>

					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Report Type</font></td>
						<td class="inputvalue" width="70%">
							<input class="image" type="radio" name="REP_TYPE" value="1" <%= checked1 %>>Daily
							<input class="image" type="radio" name="REP_TYPE" value="2" <%= checked2 %>>Weekly
							<input class="image" type="radio" name="REP_TYPE" value="3" <%= checked3 %>>Monthly
							<% small_help "Report Type" %></td>
					</tr>
		
					<tr bgcolor='#FFFFFF'>
						<td colspan=3 align=center>
							<input class=buttons type="submit" name="Generate_Report" value="Get Affiliate Report"><br>&nbsp;</font></td>
					</tr>
				</table>
				</td>
		</tr>

		<%' FOR EACH AFFILIATE, COMPUTE NUMBER OF CLICKS / ORDERS  %>
		<% rs_Store.open sql_select_affiliates,conn_store,1,1  %>
		<% if not rs_Store.eof then %>
		<tr bgcolor='#FFFFFF'>
				<td width="100%" colspan="3" height="13">
				<table border="1" width="100%" height="1" cellspacing="2" cellpadding="2" class="list">
					<tr bgcolor='#FFFFFF'>
						<td rowspan="2" align="center"><font size="1"><b>Period</b></font></td>
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
									><font size="1"><b><%= rs_store("Contact_Name") %><br>(<%= rs_store("Code") %>)</b></font></td>
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
								<td align="center" colspan="4" bgcolor="#AAAAA"><font size="1"><b>Totals</b></font></td>
							<% End If %>
					</tr>
					
					<tr bgcolor='#FFFFFF'>
						<% for i=1 to tot_affs %>
							<% If i mod 2 = 0 then %>
								<% bgcolor = "bgcolor=""#DDDDD""" %>
							<% Else %>
								<% bgcolor ="" %>
							<% End If %>
							<td align="center" <%= bgcolor %>><font size="1"><b>Clicks</b></font></td>
							<td align="center" <%= bgcolor %>><font size="1"><b>Orders</b></font></td>
							<td align="center" <%= bgcolor %>><font size="1"><b>Order Totals</b></font></td>
							<td align="center" <%= bgcolor %>><font size="1"><b>Payout</b></font></td>
						<% next %>
						<% If use_ltotals then %>
							<td align="center" bgcolor="#AAAAA"><font size="1"><b>Clicks</b></font></td>
							<td align="center" bgcolor="#AAAAA"><font size="1"><b>Orders</b></font></td>
							<td align="center" bgcolor="#AAAAA"><font size="1"><b>Order Totals</b></font></td>
							<td align="center" bgcolor="#AAAAA"><font size="1"><b>Payout</b></font></td>
						
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
						<td align="center"><font size="1">
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
							 sql_select_c1 = "select count(SS.store_id) as totClicks from affiliate_click_report SS, store_affiliates SA where SA.Affiliate_ID="&af_codes(i-1)&" and SS.CAME_FROM=SA.Code and SS.SYS_Created>='"&prev_cur_date&"' and SS.SYS_Created<'"&cur_date&"' and SS.store_id="&store_id&" and SA.Store_ID="&store_id

							 rs_Store.open sql_select_c1,conn_store,1,1	
							 If rs_Store("totClicks")="" or isnull(rs_Store("totClicks")) then
								 c_clicks = 0 
							 Else 
								 c_clicks = rs_Store("totClicks") 
							 End If 
							 rs_Store.close 
							 tot_clicks(i-1) = tot_clicks(i-1) + c_clicks 
							 l_clicks = l_clicks + c_clicks 

							 sql_select_c2 = "select count(SP.store_id) as totOrd, sum(SP.Total) as totAm from Store_Affiliates SA WITH (NOLOCK), Store_Purchases SP WITH (NOLOCK) where SP.verified=1 and SA.Affiliate_ID="&af_codes(i-1)&" and SP.CAME_FROM=SA.Code and SP.SYS_Created>='"&prev_cur_date&"' and SP.SYS_Created<'"&cur_date&"' and SA.store_id="&store_id&" and SP.store_id="&Store_id
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

							<td align="center" <%= bgcolor %>><font size="1"><%= c_clicks %></font></td>
							<td align="center" <%= bgcolor %>><font size="1"><%= c_ords %></font></td>
							<td align="center" <%= bgcolor %>><font size="1"><%= Store_Currency %><%= formatnumber(c_vals,2) %></font></td>
							<td align="center" <%= bgcolor %>><font size="1"><%= Store_Currency %><%= formatnumber(c_payout,2) %></font></td>
						<% next %>
			
						<% If use_ltotals then %>
							<td align="center" bgcolor="#AAAAAA"><font size="1"><%= l_clicks %></font></td>
							<td align="center" bgcolor="#AAAAAA"><font size="1"><%= l_ords %></font></td>
							<td align="center" bgcolor="#AAAAAA"><font size="1"><%= Store_Currency %><%= formatnumber(l_vals,2) %></font></td>
							<td align="center" bgcolor="#AAAAAA"><font size="1"><%= Store_Currency %><%= formatnumber(l_payout,2) %></font></td>
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
						<td align="center"><font size="1"><b><br>Totals<br>&nbsp;</b></font></td>
							<% gt_clicks = 0 %>
							<% gt_ords = 0 %>
							<% gt_vals = 0 %>
							<% gt_payout = 0 %>
							<% for i=1 to tot_affs %>
								<% gt_clicks = gt_clicks + tot_clicks(i-1) %>
								<% gt_ords = gt_ords + tot_ords(i-1) %>
								<% gt_vals = gt_vals + tot_vals(i-1) %>
								<% gt_payout = gt_payout + tot_payout(i-1) %>
								<td align="center"><font size="1"><b><%= tot_clicks(i-1) %></b></font></td>
								<td align="center"><font size="1"><b><%= tot_ords(i-1) %></b></font></td>
								<td align="center"><font size="1"><b><%= Store_Currency %><%= formatnumber(tot_vals(i-1),2) %></b></font></td>
								<td align="center"><font size="1"><b><%= Store_Currency %><%= formatnumber(tot_payout(i-1),2) %></b></font></td>
							<% next %>
							<% If use_ltotals then %>
								<td align="center" bgcolor="#AAAAAA"><font size="1"><b><%= gt_clicks %></b></font></td>
								<td align="center" bgcolor="#AAAAAA"><font size="1"><b><%= gt_ords %></b></font></td>
								<td align="center" bgcolor="#AAAAAA"><font size="1"><b><%= Store_Currency %><%= formatnumber(gt_vals,2) %></b></font></td>
								<td align="center" bgcolor="#AAAAAA"><font size="1"><b><%= Store_Currency %><%= formatnumber(gt_payout,2) %></b></font></td>
							<% End If %>
					</tr>

					<tr bgcolor="#DDDDDD">
						<td align="center"><font size="1"><b><br>More Details<br>&nbsp;</b></font></td>
						<% for i=1 to tot_affs %>
							<td align="center" colspan="4"><font size="1"><b><a class="link" href="JavaScript:goDetails('<%= af_codes(i-1) %>');">Click and Order Details</a></b></font></td>
						<% next %>
						<% If use_ltotals then %>
							<td align="center" colspan="4"><font size="1"><b>&nbsp;</b></font></td>
						<% End If %>
					</tr>

	</table></td></tr>
	<% else %>
		<% rs_Store.close %>
	<% End If %>
	<% createFoot thisRedirect, 0%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");
 frmvalidator.addValidation("Afiliates","req","Please choose which affiliates to run the report for.");
</script>
<% end if %>


