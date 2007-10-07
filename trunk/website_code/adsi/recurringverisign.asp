<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include virtual="common/cc_validation.asp"-->


<%

Const Request_POST = 1
Const Request_GET = 2
server.scripttimeout = 4800

'if request.servervariables("Remote_Addr") = request.servervariables("Local_Addr") then
if 1=1 then
	on error goto 0

	sql_select = "select sys_billing.*,store_settings.service_type,store_settings.overdue_payment, reseller_portion, store_settings.reseller_id as reseller,store_email,site_name from sys_billing inner join store_settings on sys_billing.store_id=store_settings.store_id where store_cancel is null and overdue_payment<98 and next_billing_date<getdate() order by next_billing_date "
	response.write sql_select
   set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

	sMailMessage = ""

	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
		    Term = mydata(myfields("payment_term"),rowcounter)
			 service_type=mydata(myfields("service_type"),rowcounter)
	       if Term = 1 then
					Term_Name = "monthly"
			 elseif Term = 3 then
					Term_Name = "quarterly"
			 elseif Term = 6 then
					Term_Name = "semi-annually"
			 elseif Term = 12 then
					Term_Name = "yearly"
			 end if

			if service_type=3 then
					Service = "bronze"
			elseif service_type = 5 then
					Service = "silver"
			elseif service_type =7 then
					Service = "gold"
			elseif service_type = 9 then
					Service = "platinum"
			elseif service_type = 10 then
					Service = "unlimited"
			elseif service_type = 11 then
					Service = "unlimited"
			else
				   Service=""
			end if

			Store_Id = mydata(myfields("store_id"),rowcounter)
			Site_Name = mydata(myfields("site_name"),rowcounter)
			next_billing_date = mydata(myfields("next_billing_date"),rowcounter)
			Sys_Created = mydata(myfields("sys_created"),rowcounter)
			reseller_id = mydata(myfields("reseller"),rowcounter)
			reseller_portion = mydata(myfields("reseller_portion"),rowcounter)
			Payment_Method = mydata(myfields("payment_method"),rowcounter)
			if not isnull(Payment_Method) then
			   Payment_Method = replace(Payment_Method," ","")
			end if
			Amount = mydata(myfields("amount"),rowcounter)
			Overdue_Payment = mydata(myfields("overdue_payment"),rowcounter)
			Discount_Until = mydata(myfields("discount_until"),rowcounter)
			if Date() < Discount_Until then
				Fee_Discount = mydata(myfields("fee_discount"),rowcounter)
			else
				Fee_Discount = 0
			end if
			total = FormatNumber(Amount - Fee_Discount)
			Store_Email= mydata(myfields("store_email"),rowcounter)


			process = 0
			success = 0
			sError =""
			
			if (Amount - Fee_Discount > 0) then
				if Overdue_Payment=0 or Overdue_Payment=3 or Overdue_Payment=7 or Overdue_Payment=14 then
				   process=1
				elseif Overdue_Payment>0 then
					total = FormatNumber(Amount - Fee_Discount)
					process=0
				end if
				
				if process=1 then
					if Payment_Method = "PayPal" then
						
						success=0
						
					else	
						First_name= mydata(myfields("first_name"),rowcounter)
						Last_name= mydata(myfields("last_name"),rowcounter)
						Company= mydata(myfields("company"),rowcounter)
						if isNull(Company) then
							Company = ""
						end if
						Address= mydata(myfields("address"),rowcounter)
						City= mydata(myfields("city"),rowcounter)
						State= mydata(myfields("state"),rowcounter)
						Zip= mydata(myfields("zip"),rowcounter)
						Country= mydata(myfields("country"),rowcounter)
						Phone= mydata(myfields("phone"),rowcounter)
						Fax= mydata(myfields("fax"),rowcounter)
						if isNull(Fax) then
							Fax = ""
						end if
						Card_Number=decrypt(mydata(myfields("card_number"),rowcounter))
						processtype = "auth_capture"
						success=process_transaction
					end if	
				end If
				if success=1 then
					sql_insert = "Insert into Sys_Payments (Store_Id,Amount,Auth_Number,AVS_Result, Card_Verif, "&_
						"Transaction_ID,Card_Ending, Payment_Description,Payment_Type,Payment_Term) values ("&Store_Id&","&total&",'"&AuthNumber&"','"&avs_result&"','"&_
						card_verif&"','"&Verified_Ref&"','"&Right(Card_Number,4)&"','"&Term_Name&" "&Service&"','Normal',"&Term&")"
					conn_store.Execute sql_insert

					sql_reseller = "Put_Reseller_Customer '"&reseller_id&"','"&Store_Id&"','"&Term&"','"&reseller_portion&"','"&term_name&"','"&now()&"','"&service&"'"
					response.write sql_reseller & "<BR>"
               conn_store.Execute sql_reseller
					
               sMailMessage = sMailMessage & chr(13) & chr(10) & Store_Id & " charge SUCCESSFULL $"& total
					

				elseif success=0 then
					Overdue_Payment = Overdue_Payment + 1
					sMessage = "Dear Customer,"&vbcrlf&vbcrlf&_
                                        "Your payment for ecommerce store, http://"&Site_Name&" (Store ID: "&Store_Id&"),"&_
					" was rejected by our payment processor on "&now()&".  This could be for many reasons including your account is over its credit limit, "&_
  						"your credit card has expired, you have cancelled the card, it was reported stolen, maintenance at your bank, etc.	If StoreSecured does not receive your "&_
  						"payment 1 week after it was originally due your store is subject to closure."&vbcrlf&vbcrlf&_
  						"To ensure uninterrupted service please go to http://manage.storesecured.com at your "&_
  						"earliest convenience and once you have logged in you may submit payment or call 1-866-324-2764 and reference your store id : " & Store_Id &" to update your billing info or update your Paypal account with the appropriate information."&vbcrlf&vbcrlf&_
  						"If you no longer wish to keep your store please login and select the cancel store link to stop all future billings and remove the store. "&_
						"As long as your store is active we will continue to try and process this charge to renew the store."&vbcrlf&vbcrlf&_
						"Sincerely,"&vbcrlf&vbcrlf&"The StoreSecured Staff"
  					 Send_Mail sNoReply_email,Store_Email,"Billing Failed for Store Id "&Store_Id& "("&Site_Name&")",sMessage
					 sMailMessage = sMailMessage & chr(13) & chr(10) & Store_Id & Email & " charge FAILED $"& total & vbcrlf& sError&vbcrlf&"overdue="&Overdue_Payment
					sql_update = "update store_settings set overdue_payment="&overdue_payment&", custom_amount="&total&", custom_description='Manual Payment' where store_id="&Store_Id
					conn_store.execute sql_update
				elseif success=2 then
					sMailMessage = sMailMessage & chr(13) & chr(10) & Store_Id & " charge PAYPAL $"& total&" "&Email
				end if
			end if

		Next
	End If
	Set Echo = Nothing
	Send_Mail sNoReply_email,"overnight@storesecured.com","Billing",sMailMessage

end if

function process_transaction()
	Exp_Month=mydata(myfields("exp_month"),rowcounter)
	Exp_Year=mydata(myfields("exp_year"),rowcounter)
	service_type=mydata(myfields("service_type"),rowcounter)
	if Payment_Method<>"eCheck" then
		if isNull(Card_Number) then
			sError=" No CC Number"
		else
		        Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")
                        parmList = "TRXTYPE=S"
                        parmList = parmList + "&ACCT="&Card_Number
                        'parmList = parmList + "&ACCT=4111111111111111"
                        parmList = parmList + "&PWD=ankle237"
                        parmList = parmList + "&USER=blac6789"
                        parmList = parmList + "&VENDOR="
                        parmList = parmList + "&PARTNER=StoreSecured"
                        parmList = parmList + "&EXPDATE="&replace(Exp_Month&Exp_Year," ","")
                        parmList = parmList + "&AMT="&formatnumber(cdbl(total),2)
                        parmList = parmList + "&STREET="&replace(Address,"&#39;","")
                        parmList = parmList + "&EMAIL="&replace(Store_Email,"&#39;","")
                        parmList = parmList + "&ZIP="&zip
                        parmList = parmList + "&FIRSTNAME="&replace(first_name,"&#39;","")
                        parmList = parmList + "&LASTNAME="&replace(last_name,"&#39;","")
                        parmList = parmList + "&TENDER=C"
                        parmList = parmList + "&PONUM="&Store_Id
                        parmList = parmList + "&SHIPTOZIP="&zip
                        parmList = parmList + "&TAXAMT=0"
                        parmList = parmList + "&TAXEXEMPT=N"
                        parmList = parmList + "&CITY="&City
                        parmList = parmList + "&CUSTCODE="&Store_Id
                        parmList = parmList + "&STATE="&replace(State,"&#39;","")
                        parmList = parmList + "&RECURRING=Y"
                        parmList = parmList + "&FREIGHTAMT=0"
                        parmList = parmList + "&COMMENT1="&Term_Name&" "&Service

                        'for real use payflow.verisign.com
                        'for testing use test-payflow.verisign.com
                        Ctx1 = client.CreateContext("payflow.verisign.com", 443, 30, "", 0, "", "")
                        curString = client.SubmitTransaction(Ctx1, parmList, Len(parmList))
                        client.DestroyContext (Ctx1)


                        Do while Len(curString) <> 0
                      	if InStr(curString,"&") Then
                      		varString = Left(curString, InStr(curString , "&" ) -1)
                      	else
                      		varString = curString
                      	end if
                      	name = Left(varString, InStr(varString, "=" ) -1)
                      	value = Right(varString, Len(varString) - (Len(name)+1))
                      	select case name
                      		case "RESULT" 
                      			resultval = value
                      		case "RESPMSG"
                      			respMessage = value
                      		case "AUTHCODE"
                      			AuthNumber = value
                      		case "PNREF"
                      			Verified_Ref = value
                      		case "AVSADDR"
                      			avsaddr = value
                      		case "AVSZIP"
                      			avszip = value
                      		case "IAVS"
                      			iavs = value
                      		case "CVV2MATCH"
                      			card_code_verif = value
                      	end select
                      	if Len(curString) <> Len(varString) Then 
                      		curString = Right(curString, Len(curString) - (Len(varString)+1))
                      	else
                      		curString = ""
                      	end if
                      Loop
                      
                      if avsaddr = "N" and avszip = "N" then
                      	avs_result = "N"
                      elseif avsaddr = "Y" and avszip = "Y" then
                      	avs_result = "Y"
                      elseif avsaddr = "N" and avszip = "Y" then
                      	avs_result = "Z"
                      elseif avsaddr = "Y" and avszip = "N" then
                      	avs_result = "A"
                      elseif iavs = "Y" then
                      	avs_result = "G"
                      elseif iavs = "X" or avszip = "X" or avsaddr = "X" then
                      	avs_result = "S"
                      end if
                      
                      if card_code_verif = "Y" then
                      	card_verif = "M"
                      elseif card_code_verif = "N" then
                      	card_verif = "N"
                      elseif card_code_verif = "X" then
                      	card_verif = "X"
                      elseif card_code_verif = "" or not Use_CVV2 then
                      	card_verif = "P"
                      end if
                      
                      If resultval <> 0 then
                      	success=0
                        sError=respMessage
                      else
                          sError=""
			  success=1
                      End IF
                      
               end if

	else
		BankABA = decrypt(mydata(myfields("bank_aba"),rowcounter))
		BankAccount = decrypt(mydata(myfields("bank_account"),rowcounter))
		BankName = mydata(myfields("bank_name"),rowcounter)
		acct_type = mydata(myfields("acct_type"),rowcounter)
		org_type = mydata(myfields("org_type"),rowcounter)
		DrvState = mydata(myfields("license_state"),rowcounter)
		DrvNumber =mydata(myfields("license_num"),rowcounter)
		dobd = mydata(myfields("license_exp_day"),rowcounter)
		dobm = mydata(myfields("license_exp_month"),rowcounter)
		doby = mydata(myfields("license_exp_year"),rowcounter)
		CheckSerial=mydata(myfields("check_num"),rowcounter)

		if processtype = "auth_capture" then
			sTransType = "DD"
		else
			sTransType = "DV"
		end if
		if org_type = "I" then
			sType = "P"
		else
			'sType = "B"
			sType = "P"
		end if
		if acct_type = "CHECKING" then
			sType = sType & "C"
		else
			sType = sType & "S"
		end if
		Set Echo = Server.CreateObject("ECHOCom.Echo")
		Randomize ' for the counter
		'******************* Required Fields *******************************
		Echo.EchoServer="https://wwws.echo-inc.com/scripts/INR200.EXE"
		Echo.counter=Store_Id
		Echo.order_type="S"
		'Echo.merchant_echo_id="123>4682520"
		'Echo.merchant_pin="54121678"
		Echo.merchant_echo_id="858>8311486"
		Echo.merchant_pin="24342434"
		Echo.billing_ip_address=Request.ServerVariables("REMOTE_ADDR")
		Echo.merchant_email=sSales_email
		Echo.grand_total=total
		Echo.sales_tax=0
		Echo.merchant_trace_nbr=Oid

		Echo.billing_first_name =first_name
		Echo.billing_last_name=last_name
		Echo.billing_company_name=company
		Echo.billing_phone=phone
		Echo.billing_address1=address
		Echo.billing_address2=""
		Echo.billing_city=city
		Echo.billing_state=state
		Echo.billing_zip=zip
		Echo.billing_email=email
		Echo.billing_country=country
		Echo.billing_fax=fax

		Echo.transaction_type=sTransType
		Echo.ec_id_type = "DL"
		Echo.ec_account=BankAccount
		Echo.ec_account_type =sType
		Echo.ec_address1=address
		Echo.ec_address2=""
		Echo.ec_bank_name=BankName
		Echo.ec_city=city
		Echo.ec_email=EMail
		Echo.ec_first_name=first_name
		Echo.ec_id_country="US"
		Echo.ec_id_exp_mm=dobm
		Echo.ec_id_exp_dd=dobd
		Echo.ec_id_exp_yy=doby
		Echo.ec_id_number=DrvNumber
		Echo.ec_id_state=DrvState

		Echo.ec_last_name=last_name
		Echo.ec_payee="StoreSecured"
		Echo.ec_payment_type="WEB"
		Echo.ec_rt=BankABA
		Echo.ec_serial_number=CheckSerial
		Echo.ec_state=State
		Echo.ec_zip=zip

		'Echo.EchoDebug="F"

		Echo.product_description=Store_Id&"-"&Term
		Echo.purchase_order_number=Store_Id&"-"&datepart("m",date())&"/"&datepart("yyyy",date())

		If Echo.Submit Then
			sError = ""
			success=1
		Else
			sError = Echo.echotype2
			success=0
		End if

	end if
	process_transaction=success
end function
%>

