<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
if checkSSL=1 then
   scheckSSL=1
else
   scheckSSL=0
end if

'****************************CODE HERE TO DISPPLAY THE DEFAULT PLAN RATES**********
intResellerId = SiteResellerID

dim strget,planrate,i,planid
redim planrate(0),planid(0)

redim preserve planrate(0)
redim preserve planrate(1)
redim preserve planrate(2)
redim preserve planrate(3)
redim preserve planrate(4)
redim preserve planrate(5)

planrate(0)="9.95"
planrate(1)="14.95"
planrate(2)="19.95"
planrate(3)="29.95"
planrate(4)="39.95"
planrate(5)="67"
'******************************CODE ENDS HERE *****************************************



'Code added by Sudha Ghogare on 13 th August 2004 to show reseller rates 
'*********************************Code starts here******************************

	if intResellerId <> "" then 
		i=1
		for i=1 to 5
					strget="Get_reseller_plan "&intResellerId&" ,"&i&""
					set rsget = conn_store.execute(strget)
					if not rsget.eof then 
						redim preserve planrate(i)
						planrate(i)=formatnumber(trim(rsget("fld_rate")),2)
					end if	
		next	
	end if	
'*********************************Code ends here******************************
	
on error goto 0

sql_select = "select Payment_Method,Payment_Term, Amount from sys_billing where store_id="&Store_Id
rs_store.open sql_select, conn_store, 1, 1
if not rs_store.eof then
	Payment_Term = rs_Store("Payment_Term")
	Payment_Method = replace(rs_Store("Payment_Method")," ","")
	Amount = rs_Store("Amount")

end if
rs_store.close

if Service_Type=1 then
	sThisPlanRate=0
elseif Service_Type=3 then
	sThisPlanRate=1
elseif Service_Type=5 then
	sThisPlanRate=2
elseif Service_Type=7 then
	sThisPlanRate=3
elseif Service_Type=9 then
	sThisPlanRate=4
else
	sThisPlanRate=5
end if

if cdbl(planrate(sThisPlanRate)) < cdbl(amount) then
	planrate(sThisPlanRate) = Amount / Payment_Term
	if Payment_term = 3 then
     	planrate(sThisPlanRate)=planrate(sThisPlanRate)/.95
     elseif Payment_term = 6 then
     	planrate(sThisPlanRate)=planrate(sThisPlanRate)/.90
     elseif Payment_term = 12 then
     	planrate(sThisPlanRate)=planrate(sThisPlanRate)/.75
     end if
end if

if scheckSSL=1 then
   sFormAction = "https://"&request.servervariables("HTTP_HOST")&"/billing_info.asp"
   'sFormAction = "billing_info.asp"
else
   sFormAction = "billing_info.asp"
end if

sFormName = "payment"
sTitle = "Upgrade Service"
sFullTitle = "My Account > Payments > Upgrade Service"
thisRedirect = "billing_info.asp"
sMenu = "account"
sSubmitName = "submit"
createHead thisRedirect %>


	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Type</B></TD><TD class="inputvalue"><select name="payment_method" size="1">
			 <% if Payment_Method <> "" then %>
			 <option selected value="<%= Payment_Method %>"><%= Payment_Method %></option>
			 <% end if %>
			 <option value="Visa">Visa</option>
			 <option value="Mastercard">Mastercard</option>
			 <option value="American Express">American Express</option>
			 <option value="Discover">Discover</option>
			 <option value="Paypal">PayPal</option>

			 </select></td><td class=inputvalue>
			 <img src=images/icon_visa.gif><img src=images/icon_mastercard.gif><img src=images/icon_discover.gif><img src=images/icon_amex.gif><BR>OR <B>PayPal</b>

	 <% small_help "Type" %></TD>
	 </TR>

	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Service</B></TD><TD class="inputvalue"><select name="service" size="1">
	 <% if Service_Type = 1 then %>
			<option value="pearl">Pearl $<%=planrate(0)%> Month (no ecommerce)</option>
	 <% end if %>
    <% if Service_Type = 3 or Trial_version=1 then %>
			<option value="bronze" selected>Bronze $<%=planrate(1)%> Month</option>
	 <% end if %>
	 <% if Service_Type = 5 or Trial_version=1 then %>
			<option value="silver">Silver $<%=planrate(2)%> Month</option>
	 <% end if %>
	 <% if Service_Type = 7 or Trial_version=1 then %>
			<option value="gold">Gold $<%=planrate(3)%> Month</option>
	 <% end if %>
	 <% if Service_Type = 9 or Trial_version=1 then %>
			<option value="platinum">Platinum $<%=planrate(4)%> Month</option>	  
	 <% end if %>
	 <% if (Amount=planrate(1) or Amount=planrate(2) or Amount = planrate(3)) and Service_Type=10 then %>
			<option value="platinum1">Platinum $15 Month</option></select></td><td class=inputvalue>&nbsp;
	 <% else %>	
			<option value="unlimited">Unlimited $<%=planrate(5)%> Month</option>
	 
			</select></td><td class="inputvalue">
	 <% end if %>
	 <% small_help "Service" %></TD>
	 </TR>

	 <TR bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Payment Term</B></TD><TD class="inputvalue" colspan=2><select name="term" size="1">
		 <%
       if Payment_Term <= 1 then %>
		 <option value="1">Monthly</option>
		 <% end if %>
		 <% if Payment_Term <= 3 then %>
		 <option value="3">Quarterly (Save 5%)</option>
		 <% end if %>
		 <% if Payment_Term <= 6 then %>
		 <option value="6">Semi-Annually (Save 10%)</option>
		 <% end if %>
		 <% if Payment_Term <= 12 then %>
		 <option selected value="12">Yearly (Save 25%)</option>
		 <% end if %>
		 </select>
	 <% small_help "Payment Term" %></TD>
	 </TR>
	 
	 <%
	 if Service_Type = 0 or Trial_Version = -1 then
		 sType = "Normal"
	 else
		 sType = "Upgrade"
	 end if
	 
	 if stype="Upgrade" then
sql_select = "select sys_billing.amount,sys_payments.transdate,sys_payments.payment_term from sys_payments inner join sys_billing on sys_payments.store_id = sys_billing.store_id where sys_payments.store_id="&Store_Id&" and sys_payments.Payment_Type='Normal' order by sys_payments.transdate desc"

        rs_store.open sql_select, conn_store, 1, 1
	if not rs_store.eof then
		Amount = rs_store("Amount")
		Sys_Created = rs_Store("transdate")
		sterm = rs_Store("Payment_Term")
	else
		 stype = "Normal"
	end if
	rs_Store.close

	if stype = "Upgrade" then
		if sterm = 12 then
			DaysUsed = DateDiff("d",Sys_Created,Date())
			DaysinPeriod = 365
			iPercent = (365-DaysUsed)/365.00
		else
			DaysUsed = DateDiff("d",Sys_Created,Date())
			DaysinPeriod = 30 * sterm
			iPercent = (DaysinPeriod - DaysUsed)/DaysinPeriod

		end if
		CreditLeft = iPercent * Amount
		OldTotal = formatNumber(total,2)
		total = (total - CreditLeft)


   end if
	end if%>
    <TR bgcolor='#FFFFFF'><TD colspan=4 align=center><input type="submit" value="Continue" name=<%= sSubmitName %>></td></tr>
	 <TR bgcolor='#FFFFFF'><TD colspan=4 class=instructions>You may upgrade and/or change your payment term of type at anytime.
	 <BR>Any unused amounts from previous payments will be applied as a credit towards new service.
   <BR><BR>
   <% if CreditLeft<>0 then %>
   You currently have an unused credit of $<%=formatNumber(CreditLeft,2) %> which will be applied to your upgrade.
   <% end if %>


	 <input type=hidden name=Special value="<%= request.querystring("Special") %>">
	 <input type=hidden name=Type value="<%= sType %>">
	 
<% if scheckSSL=0 then %>
      <table bgcolor=Red width=100%><TR bgcolor='#FFFFFF'><td class=error><B>WARNING: Your browser does not support secure pages</b><BR>We are unable to transfer you to a secure connection to collect your sensitive payment information because either your browser does not support a secure connection or you have disabled it.  You may want to choose a new browser or enable SSL before continuing.  You may also call our toll-free number 1-866-324-2764 to upgrade by phone.</b></td></tr></table>
<% end if %>
<table align=center><TR bgcolor='#FFFFFF'><td>
<!-- START OF HACKER FREE SECURITY TAG -->
<a target="_blank" href="https://secure.xentinelsecurity.com/Certificates/Verify.aspx?ref=<%= hackerfree_url %>">
<img width =139 height =40 src="//secure.xentinelsecurity.com/Certificates/sealserver.aspx?ref=<%= hackerfree_url %>&Seal=1"
 OnError = "this.onerror=null;this.src='';this.width=0;this.height=0"
 border = "0" name = "hs_seal" alt="HACKER FREE certified websites are a secure place to shop, they are 99.9% immune to hacker attacks."></a>
<!-- END OF HACKER FREE SECURITY TAG -->
</td></tr></table>

<% createFoot thisRedirect,2 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
<script src="https://ssl.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-1343888-1";
_udn="none";
_ulink=1;
urchinTracker();
</script>

