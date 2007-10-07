<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/cc_validation.asp"-->

<%
'ERROR CHECKING
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%

else

	first_name=checkStringForQ(request.form("first"))
	last_name=checkStringForQ(request.form("last"))
	address=checkStringForQ(request.form("address"))
	city=checkStringForQ(request.form("city"))
	state=checkStringForQ(request.form("state"))
	zip=checkStringForQ(request.form("zip"))
	country=checkStringForQ(request.form("country"))
	phone=checkStringForQ(request.form("phone"))
	fax=checkStringForQ(request.form("fax"))
	company=checkStringForQ(request.form("company"))
	CardCode=checkStringForQ(request.form("CardCode"))
	counter=checkStringForQ(request.form("counter"))
	email=checkStringForQ(request.form("email"))

	cc_num = checkStringForQ(request.form("cc_num"))
	mm=request.form("mm")
	yy=request.form("yy")
	email = checkStringForQ(request.form("email"))
	payment_method= checkStringForQ(request.form("payment_method"))

	term = request.form("term")
	service = request.form("service")

		if service = "silver" then
			total = FormatNumber(request.form("Cost_To_UpgradeSilver"),2)
			if term="monthly" then
				new_total = 15
			else
				new_total=135
			end if
			level = 5
		elseif service="gold" then
			total = FormatNumber(request.form("Cost_To_UpgradeGold"),2)
			level = 7
			if term="monthly" then
				new_total = 25
			else
				new_total=225
			end if
		else
			total = FormatNumber(request.form("Cost_To_UpgradePlatinum"),2)
			level = 10
			if term="monthly" then
				new_total = 40
			else
				new_total=360
			end if
		end if

	if not IsCreditCard(payment_method,cc_num) then
		Response.redirect "error.asp?Message_id=45"
	end if
	if 2000+cint(yy)=year(now) then
		if cint(mm)<month(now) then
			Response.redirect "error.asp?Message_id=46"
		end if
	end if

	processtype = "auth_capture"
	sDescription = Store_Id & "-upgrade-" & service
	sPONum = Store_Id&"-"&datepart("m",date())&"/"&datepart("yyyy",date())
		%>
		<!--#include file="process_cc.asp"-->
		<%
		  sql_insert = "Insert into Sys_Payments (Store_Id,Amount,Auth_Number,AVS_Result, Card_Verif, "&_
				"Transaction_ID) values ("&Store_Id&","&total&",'"&AuthNumber&"','"&avs_result&"','"&_
				card_verif&"','"&Verified_Ref&"')"
		  conn_store.Execute sql_insert

		  sql_insert = "Update Sys_Billing set Amount="&new_total&_
				",Term='"&term&"',First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
				"Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
				"Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
				"Card_Number='"&encrypt(cc_num)&"',Exp_Month='"&mm&"',Exp_Year='"&yy&"' "&_
				",Email='"&Email&"' where Store_Id = "&Store_ID
				
		  sql_update = "update store_settings set Trial_Version=0,Service_Type="&level&" where Store_id ="&Store_id
		  conn_store.Execute sql_update
		  Send_Mail sReport_email,sReport_email,"Store Paid Upgrade","Store " & Store_Id & " has paid " & total & " for easystorecreator "& service & " upgrade in service."





	conn_store.Execute sql_insert

	response.redirect "old_billing_info.asp?Accepted=Yes"
end if

%>
