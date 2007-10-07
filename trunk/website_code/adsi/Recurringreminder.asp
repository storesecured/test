<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include virtual="common/cc_validation.asp"-->


<%
server.scripttimeout = 4800

if request.servervariables("Remote_Addr") = request.servervariables("Local_Addr") then
'if 1=1 then
	today=now()
	threedays=formatdatetime(dateAdd("d",3,today),2)
	fourdays=formatdatetime(dateAdd("d",4,today),2)
	twentyninedays=formatdatetime(dateAdd("d",29,today),2)
	thirtydays=formatdatetime(dateAdd("d",30,today),2)


	'select those customers who are 3 days from billing or 30 days from billing for yearly payments
	sql_select = "select sys_billing.store_id,next_billing_date,payment_term,amount,payment_method,card_number,email,site_name,service_type,payment_term from Sys_Billing inner join store_settings on store_settings.store_id=sys_billing.store_id where (next_billing_date>='"&threedays&"' and next_billing_date<'"&fourdays&"') or (payment_term=12 and next_billing_date>='"&twentyninedays&"' and next_billing_date<'"&thirtydays&"') and store_cancel is null and overdue_payment<>100"
	response.write sql_select
   set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

	sMailMessage = ""

	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			Store_Id = mydata(myfields("store_id"),rowcounter)
			next_billing_date=formatdatetime(mydata(myfields("next_billing_date"),rowcounter),2)
			Store_Email = mydata(myfields("email"),rowcounter)
			Site_Name = mydata(myfields("site_name"),rowcounter)
			Payment_Method = trim(mydata(myfields("payment_method"),rowcounter))
			Service_type = mydata(myfields("service_type"),rowcounter)
			Payment_Term = mydata(myfields("payment_term"),rowcounter)
			Card_Number = decrypt(mydata(myfields("card_number"),rowcounter))
			Amount = mydata(myfields("amount"),rowcounter)

			if Payment_Term=1 then
				service_name="monthly"
			elseif Payment_Term=3 then
				service_name="quarterly"
			elseif Payment_Term=6 then
				service_Name="semi-annual"
			elseif Payment_Term=12 then
				service_Name="yearly"
			end if

			if Service_Type=3 then
				service_name=service_name&" bronze"
			elseif Service_Type=5 then
				service_name=service_name&" silver"
			elseif Service_Type=7 then
				service_name=service_name&" gold"
			elseif Service_Type=9 then
				service_name=service_name&" platinum"
			elseif Service_Type=10 or Service_Type=11 then
				service_name=service_name&" unlimited"
			end if

			if left(Payment_Method,6)="PayPal" then
				sBilling="PayPal account"
			else
				sBilling=Payment_Method&" ending in "&right(card_number,4)
			end if
			sSubject ="Billing Reminder for Store Id "&Store_Id&" ("&Site_Name&")"
			sText = "Dear Customer,"&vbcrlf&vbcrlf&_
				"Your ecommerce store, http://"&Site_Name&" (Store ID: "&Store_Id&"),"&_
				" is coming up for automatic renewal shortly."&_
				"  Based on your current preferences your "&sBilling&" will be automatically billed $"&_
				Amount&" on "&next_billing_date&" for "&service_name&" service."&_
				vbcrlf&vbcrlf&"If you would like to change your service, billing, payment "&_
				"method or cancel service you may do so by logging into your store from "&_
				"http://manage.storesecured.com and selecting the My Account Menu at the top "&_
				"of the screen."&vbcrlf&vbcrlf&"This is a courtesy reminder and no action is required "&_
				"unless you wish to modify the above options for your stores billing."&vbcrlf&vbcrlf&_
            "To print previous invoices please visit My Account-->Prior Invoices"&vbcrlf&_
            "To update your credit card details or billing address please visit from My Account-->Update Payment."&vbcrlf&_
            "To upgrade to a new plan or billing term please visit from My Account-->Upgrade Service."&vbcrlf&_
            "To cancel your store service please vist My Account-->Cancel Store."&vbcrlf&vbcrlf&_
				"Sincerely, "&vbcrlf&vbcrlf&"The StoreSecured Staff"
			Send_Mail sNoReply_email,Store_Email,sSubject,sText
					 
		Next
	End If
end if	
%>

