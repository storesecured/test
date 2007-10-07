<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include virtual="common/crypt.asp"-->

<% 
'Code added by Devki Anote on 17 th August 2004 to show reseller rates
'*********************************Code starts here******************************

intResellerId = SiteResellerID
iResellerPercent = 1


dim strget,planrate,i,planid
redim planrate(0),planid(0)
'*************************************************************
redim preserve planrate(0)
redim preserve planrate(1)
redim preserve planrate(2)
redim preserve planrate(3)
redim preserve planrate(4)
redim preserve planrate(5)

planrate(0) = "9.95"
planrate(1) = "14.95"
planrate(2) = "19.95"
planrate(3) = "29.95"
planrate(4) = "39.95"
planrate(5) = "67"

'************************************************************
i=1

if intResellerId <> "" then 
		i=1
		for i=1 to 5
					strget="Get_reseller_plan "&intResellerId&" ,"&i&""
					set rsget = conn_store.execute(strget)
					if not rsget.eof then 
						redim preserve planrate(i)
						planrate(i)=trim(rsget("fld_rate"))
					end if	
		next
	end if	
' code here to retrive the sell rates of the easystorecreator

'To find out how many customers are there for a particular reseller.
sqlGetData= "select count(store_id) as count from store_settings where reseller_id="&intResellerId&" and trial_version=0 and service_type>0"
set rsGetData =  conn_store.execute(sqlGetData)
if not rsGetData.eof then 
	intNoCustReferred=trim(rsGetData("count"))
end if
rsGetData.close
set rsGetData=nothing

intNoCustReferred = trim(cint(intNoCustReferred))

'Depending on the user range find out userlimit id.
sqlGetData1= "select fld_plan_user_limit_id,fld_user_range from TBL_Plan_User_Limit_Master where fld_user_low<="&intNoCustReferred&" and fld_user_high>="&intNoCustReferred
set rsGetData1 =  conn_store.execute(sqlGetData1)
if not rsGetData1.eof then
	intLimitid= trim(rsGetData1("fld_plan_user_limit_id"))
end if
rsGetData1.close
set rsGetData1=nothing


'It will show all the rates of easy store creator depending on the Userlimit id.
sqlGetData2="select fld_plan_esc_id,fld_plan_name_id,fld_plan_user_limit_id,fld_rate,fld_plan_date from TBL_Plan_Esc where fld_plan_user_limit_id="&intLimitid&" order by fld_plan_Date"
set rsGetData2 =  conn_store.execute(sqlGetData2)
redim ESCplanrate(0),EScplanid(0),differance(0)
i=1
if not rsGetData2.eof then
	while not rsGetData2.eof 
		redim preserve EScplanrate(i)
		ESCplanrate(i)=trim(rsGetData2("fld_rate"))
		rsGetData2.movenext
		i = i + 1
	wend
end if


'code here to find the differance between the two rates
for i=1 to 5
	redim preserve differance(i)
	differance(i) =  planrate(i)-ESCplanrate(i)
next	
'*********************************Code ends here******************************


discount=0
if request.form("coupon_code") = "Freemerchant" then
	discount = .20
elseif request.form("coupon_code") = "25PERCENT" then
	discount = .25
elseif request.form("Special") = "cws" then
   discount = .30
elseif request.form("Special") = "2nd" then
   discount = .20
elseif request.form("Special") = "EH20" then
   discount = .20
elseif request.form("coupon_code") <> "" then
	fn_error "Invalid Coupon Entry"
end if

'RETRIVE CURRENT VALUES FROM THE DATABASE
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
	Card_Number=rs_store("Card_Number")
	Exp_Month=rs_Store("Exp_Month")
	Exp_Year=rs_Store("Exp_Year")
	Payment_Method_Current = replace(rs_Store("Payment_Method")," ","")
	Bank_Name = rs_Store("Bank_Name")
	Bank_Account = decrypt(rs_store("Bank_Account"))
	Bank_ABA = decrypt(rs_store("Bank_ABA"))
	Acct_Type = rs_Store("Acct_Type")
	Org_Type = rs_Store("Org_Type")
	Check_Num = rs_Store("Check_Num")
	License_Num = rs_Store("License_Num")
	License_State = rs_Store("License_State")
	License_Exp_Day = rs_Store("License_Exp_Day")
	License_Exp_Month = rs_Store("License_Exp_Month")
	License_Exp_Year = rs_Store("License_Exp_Year")
	Amount=rs_Store("Amount")
	Payment_term=rs_Store("Payment_term")
else

	Company= Store_Company
	Address= Store_Address1
	City= Store_City
	Zip= Store_Zip
	Country= Store_Country
	Phone= Store_Phone
	Fax= Store_Fax
	EMail= Store_Email
	Amount=0
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

if cdbl(planrate(sThisPlanRate)) < cdbl(Amount) then
	planrate(sThisPlanRate) = (Amount / Payment_Term)
	if Payment_term = 3 then
     	planrate(sThisPlanRate)=planrate(sThisPlanRate)/.95
     elseif Payment_term = 6 then
     	planrate(sThisPlanRate)=planrate(sThisPlanRate)/.90
     elseif Payment_term = 12 then
     	planrate(sThisPlanRate)=planrate(sThisPlanRate)/.75
     end if
end if

sFormAction = "process_billing.asp"
sFormName = "payment"
sTitle = "Upgrade Service Continue"
sFullTitle = "My Account > Payments > <a href=billing.asp class=white>Upgrade Service</a> > Continue"
thisRedirect = "billing_info.asp"
sMenu = "account"
addPicker=1
sSubmitName="submit"
createHead thisRedirect

payment_method = request("payment_method")
service = request("service")
term = request("term")
stype = request("type")

if payment_method = "" or service = "" or term = "" or stype = "" then
	'response.redirect "billing.asp"
end if

if service = "pearl" then
	total = planrate(0)
	level = 1
	amounttoreseller = 0
	if intResellerId<>"" then 
		ESCAmount = 0
	end if
elseif service = "bronze" then
	total = planrate(1)
	level = 3
	amounttoreseller = differance(1)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(1)
	end if	
elseif service = "silver" then
	total = planrate(2)
	level = 5
	amounttoreseller = differance(2)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(2)
	end if	
elseif service="gold" then
	total = planrate(3)
	level = 7
	amounttoreseller = differance(3)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(3)
	end if	
elseif service="platinum" then
	total = planrate(4)
	level = 9
	amounttoreseller = differance(4)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(4)
	end if	
elseif service="unlimited" or service="store" then
	total = planrate(5)
	level = 11
	amounttoreseller = differance(5)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(5)	
	end if
elseif service="platinum1" then
	total = 15
	level = 10
	amounttoreseller = differance(5)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(5)
	end if
else
	total = 67
	level = 11
	amounttoreseller = differance(5)
	if intResellerId<>"" then 
		ESCAmount = ESCplanrate(5)
	end if	
end if

if cint(term) = 12 then
	ResellerRate = total * 12 * .75
	total = total * 12 * .75
	term_name = "yearly"
	amounttoreseller = amounttoreseller * 12 * .75
	ESCAmount = ESCAmount * 12 * .75
	
elseif cint(term) = 6 then
	ResellerRate = total * 6 * .90
	total = total * 6 * .90
	term_name = "semi-annual"
	amounttoreseller = amounttoreseller * 6 * .90
	ESCAmount = ESCAmount * 6 * .90
	
elseif cint(term) = 3 then
	ResellerRate = total * 3 * .95
	total = total * 3 * .95
	term_name = "quarterly"
	amounttoreseller = amounttoreseller * 3 * .95
	ESCAmount = ESCAmount * 3 * .95
	
else
	ResellerRate = total
	total = total
	term_name = "monthly"
	amounttoreseller = amounttoreseller
	ESCAmount = ESCAmount
	
end if


CreditLeft = 0
if discount<>0 then
        oldtotal=total
	total = total * (1-discount)
	discount_amount=oldtotal-total

end if
OldTotal = formatnumber(total,2)

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
			iPercent = (365-DaysUsed)/365.00
		else
			DaysUsed = DateDiff("d",Sys_Created,Date())
			DaysinPeriod = 30 * sterm
			iPercent = (DaysinPeriod - DaysUsed)/DaysinPeriod

		end if
		CreditLeft = iPercent * Amount
		OldTotal = formatNumber(total,2)
		total = (total - CreditLeft)
		iResellerPercent = total / OldTotal
		if iResellerPercent < 0 then
		   iResellerPercent = 0
                end if



	end if

end if

'reseller gets 50% of profit.
amounttoreseller = formatNumber(amounttoreseller / 2)

'if this is an upgrade adjust amount accordingly
actualamounttoreseller = formatNumber(amounttoreseller * iResellerPercent,2)


CreditLeft = FormatNumber(CreditLeft,2)
if total < 1.5 then
	total = formatnumber(1.5,2)
else
	 total = FormatNumber(total,2)
end if

on error resume next
%>
<tr bgcolor='#FFFFFF'><TD colspan=3><input type="button" value="Change payment type" OnClick=JavaScript:self.location="billing.asp" class=buttons>
<%
if payment_method = "Paypal" or payment_method = "PayPal" then
	Return_Addr = "http://"&request.servervariables("http_host")&"/old_billing_info.asp"
	Notify_Addr = "http://"&request.servervariables("http_host")&"/paypal_notify2.asp"
	Cancel_Addr = "http://"&request.servervariables("http_host")&"/admin_home.asp"

	if CreditLeft <> 0 then
		 grand_total = OldTotal
	 else
		 grand_total = total
	 end if
	%>
	</form>

	<form method='POST' action='https://www.paypal.com/cgi-bin/webscr' name='Payment'>
	<input type='hidden' name='business' value='<%=sPaypal_email %>'>
	<input type="hidden" name="cmd" value="_xclick-subscriptions">
	<input type="hidden" name="redirect_cmd" value="_xclick">
	<input type="hidden" name="bn" value="EasyStoreCreator.EasyStoreCreator">
	<tr bgcolor='#FFFFFF'><td colspan=3>Click on the button below to pay using Paypal
	<% if Service_Type > 0 and (Payment_Method_Current="Paypal" or Payment_Method_Current = "PayPal") then %>
	<BR><BR>Once your new Paypal subscription is successfully created please cancel your previous subscription to ensure that your store is not billed twice.
	<% end if %>
	<BR><BR>
  </td></tr>

	<tr bgcolor='#FFFFFF'><td colspan=3 bgcolor=yellow><B>Important:</b> Please note that payments made using <B>Paypals eCheck</b> service are not recorded until the payment has been marked as cleared by PayPal, this can take up to 5 days.  If you wish to start using the upgraded features immediately you must pay using Paypal Instant Transfer, Paypal Balance, Paypal credit card or one of our other payment methods.<BR></td></tr>
	 

	<% if CreditLeft <> 0 then %>
	 <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Total</B></TD><TD class="inputvalue">$<%= OldTotal %>
	 <% small_help "Total" %></TD>
	 </TR>
	 <tr bgcolor='#FFFFFF'>
	 <% if CreditLeft < 0 then %>
	 <TD class="inputname"><B>Overdue Charges</B></TD><TD class="inputvalue"><font color=red>$<%= FormatNumber(CreditLeft * -1,2) %></font>
	 <% else %>
	 <TD class="inputname"><B>Unused Credit</B></TD><TD class="inputvalue"><font color=red>$<%= CreditLeft %></font>
	 <% end if %>
	 <% small_help "Credit Left" %></TD>
	 </TR>
	 <input type="hidden" name="a1" value="<%= total  %>">
	 <input type="hidden" name="p1" value="<%= term %>">
	 <input type="hidden" name="t1" value="M">
	 <% end if %>
	 <% if discount then %>
	 <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Discount</B></TD><TD class="inputvalue"><font color=red>$<%= formatnumber(discount_amount) %></font>

	 <% small_help "Discount" %></TD>
	 </TR>
	 <% end if %>
	 <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Amount Due</B></TD><TD class="inputvalue">$<%= Total %>&nbsp;<%= service %>&nbsp;<%= term_name %>

	 <% small_help "Charge Total" %></TD>
	 </TR>

	<tr bgcolor='#FFFFFF'><td colspan=3 align=center>
	<input type="hidden" name="item_name" value="<%= service %>&nbsp;<%= term_name %> service">

	<% if CreditLeft > 0 then %>
	<input type="hidden" name="a3" value="<%= OldTotal %>">
	<% else %>
	<input type="hidden" name="a3" value="<%= total %>">
	<% end if %>
	<input type="hidden" name="p3" value="<%= term %>">
	<input type="hidden" name="t3" value="M">
	<input type="hidden" name="src" value="1">
	<input type="hidden" name="sra" value="1">
	 	<!--Code added for second reseller amount -->
	 <input type="hidden" name="hidResellerAmt" value="<%= amounttoreseller%>">
	 <input type="hidden" name="hidResellerAmt2" value="<%= actualamounttoreseller%>">
	 <input type="hidden" name="ESCAmount" value="<%=ESCAmount%>">
	 <input type="hidden" name="ResellerRate" value="<%=ResellerRate%>">
	 
	 <!-- Code added for customer amount-->
	 
	<input type="hidden" name="notify_url" value="<%= Notify_Addr %>?Level=<%= level %>&Term=<%= term %>&Term_Name=<%= term_name %>&Service=<%= Service %>&Type=Normal&Grand_Total=<%= Grand_Total %>">


<!--#include file="capture_billing.asp"-->
<% put_paypal
createFoot thisRedirect,0

else

	put_normal
	
	IPAddress = Request.ServerVariables("REMOTE_ADDR")
  Set ipObj = Server.CreateObject("IP2Location.Country")
  ' initialize IP2Location™ Component
  If ipObj.Initialize("melanie@easystorecreator.com-2005-7-pQ3GpeMWWo1m") <> "OK" Then
  'initialization failed
  End If
  ' it will return "US"
  CountryName = ipObj.LookUpFullName(IPAddress)
  Set ipObj = nothing
%>
	 <% if CreditLeft <> 0 then %>
	 <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Total</B></TD><TD class="inputvalue">$<%= OldTotal %>
	 <% small_help "Total" %></TD>
	 </TR>
	 <%
	 if CreditLeft < 0 then
			  CreditLeft = (CreditLeft * -1)
			  sTitle = "Overdue" %>
			  <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B><%= sTitle %></B></TD><TD class="inputvalue"><font color=red>$<%= CreditLeft %></font>
	 <% small_help "Credit Left" %></TD>
	 </TR>
	 <% else
		  sTitle = "Credit Left" %>
		  <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B><%= sTitle %></B></TD><TD class="inputvalue"><font color=red>-$<%= CreditLeft %></font>
	 <% small_help "Credit Left" %></TD>
	 </TR>
	 <% end if %>
	 <% end if %>
	 <% if discount then %>
	 <tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Discount</B></TD><TD class="inputvalue"><font color=red>$<%= formatnumber(discount_amount) %></font>

	 <% small_help "Discount" %></TD>
	 </TR>
	 <% end if %>
	<tr bgcolor='#FFFFFF'>
	 <TD class="inputname"><B>Amount Due</B></TD><TD class="inputvalue">$<%= Total %>&nbsp;<%= service %>&nbsp;<%= term_name %>

	<% if total <= 1.5 then %>
	<BR><font color=red>($1.50 minimum charge amount)</font>
	 <% end if %>
	 <% small_help "Charge Total" %></TD>
	 </TR>


	 <input type=hidden name=counter value="<%= DatePart("d", Date()) %><%= Store_Id %>">
	 <input type=hidden name=total value="<%= total %>">
	 <input type=hidden name=billing_type value="Normal">
	 <% if CreditLeft <> 0 then %>
		 <% grand_total = OldTotal%>
		 <input type=hidden name=grand_total value="<%= OldTotal %>">
	 <% else %>
		 <% grand_total = total %>
		 <input type=hidden name=grand_total value="<%= total %>">
	 <% end if %>
	 
	 	<!--Code added for reseller amount -->
	 <input type="hidden" name="hidResellerAmt" value="<%= amounttoreseller%>">
	 <input type="hidden" name="hidResellerAmt2" value="<%= actualamounttoreseller%>">
	 <input type="hidden" name="ESCAmount" value="<%=ESCAmount%>">
	 <input type="hidden" name="ResellerRate" value="<%=ResellerRate%>">
	 <!-- Code added for customer amount-->
	 <input type=hidden name=term value="<%= term %>">
	 <input type=hidden name=term_name value="<%= term_name %>">
	 <input type=hidden name=service value="<%= service %>">
	 <input type=hidden name=level value="<%= level %>">
	 <input type=hidden name=description value="<%=Store_Id %>-<%= level %>-<%= term %>">
	 <input type=hidden name=ponum value="<%= Store_Id %>-<%= datepart("m",date())%>/<%= datepart("yyyy",date())%>-<%= level %>-<%= term %>">
	 <tr bgcolor='#FFFFFF'><TD colspan=3 align=center><font size=1>Service can be cancelled at anytime by filling out the cancellation request form.
    <% if payment_method="eCheck" then %>
       <BR>*Please note that returned checks are subject to a $25 collection fee.
   <% end if %></font></td></tr>
	 <tr bgcolor='#FFFFFF'><TD colspan=3 align=center><input type="submit" value="Pay Now" name=<%= sSubmitName %> class=buttons></td></tr>

   <tr bgcolor='#FFFFFF'><TD colspan=3 align=center><input type=hidden name=ip_address value=<%= IPAddress %>><input type=hidden name=ip_country value="<%= CountryName %>">
   <B>Charges will show on your statement as <I>StoreSecured, Inc.</i></b><BR><BR>
   Your IP Address (<%= IPAddress %>) has been recorded and will be provided to the necessary authorities in the event of unauthorized or fraudulent charges.
   <table align=center><tr bgcolor='#FFFFFF'><td align=center>
<!-- START OF HACKER FREE SECURITY TAG -->
<!-- Secure ID - Do Not Remove! --><div id="digicertsitesealcode" style="width: 81px; margin: 10px auto 10px 10px;" align="center"><script language="javascript" type="text/javascript" src="https://www.digicert.com/custsupport/sealtable.php?order_id=<%=ssl_oid%>&seal_type=a&seal_size=large"></script><script language="javascript" type="text/javascript">coderz();</script></div><!-- END Secure ID -->
<BR><BR><a target="_blank" href="https://secure.xentinelsecurity.com/Certificates/Verify.aspx?ref=<%= hackerfree_url %>">
<img width =139 height =40 src="//secure.xentinelsecurity.com/Certificates/sealserver.aspx?ref=<%= hackerfree_url %>&Seal=1"
 OnError = "this.onerror=null;this.src='';this.width=0;this.height=0"
 border = "0" name = "hs_seal" alt="HACKER FREE certified websites are a secure place to shop, they are 99.9% immune to hacker attacks."></a>
<!-- END OF HACKER FREE SECURITY TAG -->

</td></tr></table>
</td></tr>
<%   createFoot thisRedirect,2

if payment_method <> "Paypal" and payment_method <> "PayPal" then %>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("first","req","Please enter your first name.");
 frmvalidator.addValidation("last","req","Please enter your last name.");
 frmvalidator.addValidation("address","req","Please enter your address");
 frmvalidator.addValidation("city","req","Please enter your city");
 frmvalidator.addValidation("State","req","Please enter your state.");
 frmvalidator.addValidation("zip","req","Please enter your zip.");
 frmvalidator.addValidation("Phone","req","Please enter your phone number.");
 frmvalidator.addValidation("email","req","Please enter your email address.");
 frmvalidator.addValidation("email","email","Please enter your email address.");
<% if payment_method = "eCheck" then %>
frmvalidator.addValidation("BankName","req","Please enter your bank name.");
frmvalidator.addValidation("BankABA","req","Please enter your bank routing number.");
frmvalidator.addValidation("BankAccount","req","Please enter your bank account number.");
frmvalidator.addValidation("CheckSerial","req","Please enter your check serial number.");
frmvalidator.addValidation("DrvNumber","req","Please enter your drivers license number.");
frmvalidator.addValidation("DrvState","req","Please enter your drivers license state.");
frmvalidator.addValidation("dobd","req","Please enter your drivers license expiration.");
frmvalidator.addValidation("dobm","req","Please enter your drivers license expiration.");
frmvalidator.addValidation("doby","req","Please enter your drivers license expiration.");
<% end if %>
</script>
<script src="https://ssl.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-1343888-1";
_udn="none";
_ulink=1;
urchinTracker();
</script>
<%
end if
end if

%>


