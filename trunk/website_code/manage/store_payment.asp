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
	%><!--#include file="Include/Error_Template.asp"--><%

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
	if term <> "yearly" then
		total = FormatNumber(20,2)
	else
		total = FormatNumber(180,2)
	end if

	if not IsCreditCard(payment_method,cc_num) then
		Response.redirect "error.asp?Message_id=45"
	end if
	if 2000+cint(yy)=year(now) then
		if cint(mm)<month(now) then
			Response.redirect "error.asp?Message_id=46"
		end if
	end if
	

	sql_select = "select * from Sys_Billing where Store_Id="&Store_Id
	rs_store.open sql_select, conn_store, 1, 1
     if rs_store.eof then
		newrec="1"
	  processtype = "auth_capture"
     else
		newrec="0"
		processtype = "auth_only"
	end if
	rs_store.close

	%>
	<!--#include file="process_cc.asp"-->
	<%


 if newrec=1 then
		sql_insert = "Insert into Sys_Billing (Store_ID,Amount,Term,First_Name, "&_
			"Last_Name, Address, City, State, Zip, Country, Phone,Payment_Method, "&_
			"Card_Number, Exp_Month, Exp_Year,Email) Values ("&Store_Id&","&total&",'"&term&"','"&_
			first_name&"','"&last_name&"','"&address&"','"&city&"','"&_
			state&"','"&zip&"','"&country&"','"&phone&"','"&payment_method&"','"&_
			encrypt(cc_num)&"','"&mm&"','"&yy&"','"&Email&"')"
	else
		sql_insert = "Update Sys_Billing set Amount="&total&_
			",Term='"&term&"',First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
			"Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
			"Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
			"Card_Number='"&encrypt(cc_num)&"',Exp_Month='"&mm&"',Exp_Year='"&yy&"' "&_
			",Email='"&Email&"' where Store_Id = "&Store_ID
 end if

	conn_store.Execute sql_insert

	sql_insert = "Insert into Sys_Payments (Store_Id,Amount,Auth_Number,AVS_Result, Card_Verif, "&_
		"Transaction_ID) values ("&Store_Id&","&total&",'"&AuthNumber&"','"&avs_result&"','"&_
		card_verif&"','"&Verified_Ref&"')"
		response.write sql_insert
	conn_store.Execute sql_insert

	if newrec=1 then
		sql_update = "update store_settings set Trial_Version=0,Failed_Payment=0,Close_Account=0 where Store_id ="&Store_id
		conn_store.Execute sql_update
	end if

	Send_Mail sReport_email,sReport_email,"Store Paid","Store " & Store_Id & " has paid " & total & " for easystorecreator service to be activated."
	response.redirect "billing_info.asp?Accepted=Yes"
end if

%>
