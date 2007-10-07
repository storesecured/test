<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/creditcardfrauddetection.class"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/cc_validation.asp"-->

<%
'ERROR CHECKING
on error goto 0
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
   Error_Log = Form_Error_Handler(Request.Form)
   %><!--#include virtual="common/Error_Template.asp"--><%

else
   Billing_Type=request.form("billing_type")
   if Billing_Type="Normal" then
     sDescription = request.form("description")
     sPONum = request.form("ponum")

     term = request.form("term")
     term_name = request.form("term_name")
     service = request.form("service")
     level = request.form("level")
     grand_total = request.form("grand_total")
     Reseller_total = trim(request.form("hidResellerAmt"))
     if Reseller_total="" or not isNumeric(reseller_total) or isNull(reseller_total) then
        Reseller_Total=0
     end if
     Reseller_today = trim(request.form("hidResellerAmt2"))
     if reseller_today="" or not isNumeric(reseller_today) or isNull(reseller_today) then
        Reseller_Today=0
     end if
     HidEscAmount = trim(Request.Form("ESCAmount"))
     ResellerRate = trim(Request.Form("ResellerRate"))
     sRecurring = "Y"
   elseif Billing_Type = "Update" then

   else
     sDescription = Store_Id & "-" & service&"-"&termdesc
     sDescription = request.form("description")
     sPONum = Store_Id&"-"&datepart("m",date())&"/"&datepart("yyyy",date())
     sRecurring = "N"
   end if

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
   counter=checkStringForQ(request.form("counter"))
   email=checkStringForQ(request.form("email"))
   Reseller_total = trim(request.form("hidResellerAmt"))
     if Reseller_total="" or not isNumeric(reseller_total) or isNull(reseller_total) then
        Reseller_Total=0
     end if
     Reseller_today = trim(request.form("hidResellerAmt2"))
     if reseller_today="" or not isNumeric(reseller_today) or isNull(reseller_today) then
        Reseller_Today=0
     end if
   payment_method=checkStringForQ(request.form("payment_method"))
   processtype = "auth_capture"

   total = request.form("total")
   if isNumeric(total) then
      total = formatnumber(total,2)
   else
      Response.redirect "error.asp?Message_id=100&Message_Add="&server.urlencode("Your transaction could not be processed due to an invalid total.")
   end if

   if payment_method = "eCheck" then
      BankABA = Request.Form("BankABA")
      BankAccount = Request.Form("BankAccount")
      BankName = Request.Form("BankName")
      acct_type = Request.Form("acct_type")
      org_type = Request.Form("org_type")
      DrvState = Request.Form("DrvState")
      DrvNumber = Request.Form("DrvNumber")
      dobd = Request.Form("dobd")
      dobm = Request.Form("dobm")
      doby = Request.Form("doby")
      CheckSerial=request.form("CheckSerial")
   else
      cc_num = checkStringForQ(request.form("cc_num"))
      current_cc = request.form("current_cc")
      if cc_num = "" then
         cc_num =  decrypt(current_cc)
      end if
      mm=request.form("mm")
      yy=request.form("yy")
      CardCode=checkStringForQ(request.form("CardCode"))

      if not IsCreditCard(payment_method,cc_num) then
         Response.redirect "error.asp?Message_id=45"
      end if
      if 2000+cint(yy)=year(now) then
         if cint(mm)<month(now) then
            Response.redirect "error.asp?Message_id=46"
         end if
      end if
   end if

   if Billing_Type <> "Update" then
      %>
      <!--#include file="process_cc.asp"-->
      <%


         if payment_method="eCheck" then
           sRight = Right(BankAccount,4)
         else
           sRight = Right(cc_num,4)
         end if

         if Billing_Type="Custom" then
            sPaymentDescription = Custom_Description
            sTerm = 1
            if sPaymentDescription = "Manual Payment" then
               Billing_Type="Normal"
            end if
         else
            sPaymentDescription = Term_Name & " " & Service
            if isnumeric(Term) then
               sTerm = Term
            else
                sTerm=1
            end if
         end if

         IP_Address=request.form("IP_Address")
         IP_Country=request.form("IP_Country")

         sql_insert_payments = "Insert into Sys_Payments (Store_Id,Amount,Auth_Number,AVS_Result, Card_Verif, "&_
               "Transaction_ID,Card_Ending, Payment_Description,Payment_Type,Payment_Term,IP_Address, IP_Country, Fraud_Score, Fraud_Result) values ("&Store_Id&","&total&",'"&AuthNumber&"','"&avs_result&"','"&_
               card_verif&"','"&Verified_Ref&"','"&sRight&"','"&left(sPaymentDescription,50)&"','"&Billing_Type&"',"&sTerm&",'"&IP_Address&"','"&IP_Country&"',"&Fraud_Score&",'"&Fraud_Result&"')"

         '''======================
         
         '' code for insert into main tabel 
         select_q= "select * from Sys_Support_temp where store_id="&Store_id 
         rs_store.open select_q, conn_store, 1, 1
         do while not  rs_store.eof
			sService=sService & " " & rs_store("Subject")
			sMessage=rs_store("sMessage")
			sSubMessage= rs_store("Subject")
			sdetail = rs_store("detail")
			sql_insert = "New_Support_Request "&Store_Id&",'"&sSubMessage &"','"&sdetail&"','Customer','"&Store_Email&"'"
			conn_store.Execute sql_insert

			rs_store.movenext
         loop
         rs_store.close
         
        sql_delete = "delete  from Sys_Support_temp where store_id="&Store_id & " and status=1"
		conn_store.Execute sql_delete
		
        sql_update = "update store_settings set Custom_Amount="&request.form("total")&",custom_description='"&sService&"' where Store_id ="&Store_id
		conn_store.Execute sql_update
		
         ''''========================
         if sPaymentDescription = "Manual Payment" and request.form("Save_Info")<>"TRUE" then
             Billing_Type="Custom"
          end if

   end if

   if Billing_Type="Normal" then
     sql_select = "select Store_Id from Sys_Billing where Store_Id="&Store_Id
     rs_store.open sql_select, conn_store, 1, 1
     if rs_store.eof then
        newrec=1
     else
        newrec=0
     end if
     rs_store.close
        if newrec = 1 then
          if payment_method="eCheck" then
             sql_insert = "Insert into Sys_Billing (Store_ID,Amount,Term,First_Name, "&_
                 "Last_Name, Address, City, State, Zip, Country, Phone,Payment_Method, "&_
                 "Bank_Name, Bank_ABA, Bank_Account,Acct_Type,Org_Type,Check_Num, "&_
                  "License_Num, License_State, License_Exp_Day, License_Exp_Month, "&_
                  "License_Exp_Year, Email, Payment_Term,reseller_portion) Values ("&Store_Id&","&_
                  grand_total&",'"&term_name&"','"&_
                 first_name&"','"&last_name&"','"&address&"','"&city&"','"&_
                 state&"','"&zip&"','"&country&"','"&phone&"','"&payment_method&"','"&_
                 BankName&"','"&encrypt(BankABA)&"','"&encrypt(BankAccount)&"','"&_
                 acct_type&"','"&org_type&"','"&CheckSerial&"','"&DrvNumber&"','"&_
                 DrvState&"','"&dobd&"','"&dobm&"','"&doby&"','"&Email&"',"&term&","&Reseller_total&")"
          else
             sql_insert = "Insert into Sys_Billing (Store_ID,Amount,Term,First_Name, "&_
                 "Last_Name, Address, City, State, Zip, Country, Phone,Payment_Method, "&_
                 "Card_Number, Exp_Month, Exp_Year,Email,Payment_Term,reseller_portion) Values ("&Store_Id&","&grand_total&",'"&term_name&"','"&_
                 first_name&"','"&last_name&"','"&address&"','"&city&"','"&_
                 state&"','"&zip&"','"&country&"','"&phone&"','"&payment_method&"','"&_
                 encrypt(cc_num)&"','"&mm&"','"&yy&"','"&Email&"',"&term&","&Reseller_total&")"
          end if
     else
          if payment_method="eCheck" then
             sql_insert = "Update Sys_Billing set Amount="&grand_total&_
                 ",Term='"&term&"',First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
                 "Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
                 "Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
                 "Bank_Name='"&BankName&"',Bank_ABA='"&encrypt(BankABA)&"',Bank_Account='"&encrypt(BankAccount)&"', "&_
                 "Acct_Type='"&acct_type&"',Org_Type='"&org_type&"',Check_Num='"&CheckSerial&"', "&_
                 "License_Num='"&DrvNumber&"',License_State='"&DrvState&"',License_Exp_Day='"&dobd&"', "&_
                 "License_Exp_Month='"&dobm&"',License_Exp_Year='"&doby&"'"&_
                  ",Email='"&Email&"',Payment_Term="&Term&",Sys_Created='"&Now()&"',reseller_portion="&Reseller_total&" where Store_Id = "&Store_ID
              if sPaymentDescription = "Manual Payment" then
              
                sql_insert = "Update Sys_Billing set "&_
                 "First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
                 "Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
                 "Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
                 "Bank_Name='"&BankName&"',Bank_ABA='"&encrypt(BankABA)&"',Bank_Account='"&encrypt(BankAccount)&"', "&_
                 "Acct_Type='"&acct_type&"',Org_Type='"&org_type&"',Check_Num='"&CheckSerial&"', "&_
                 "License_Num='"&DrvNumber&"',License_State='"&DrvState&"',License_Exp_Day='"&dobd&"', "&_
                 "License_Exp_Month='"&dobm&"',License_Exp_Year='"&doby&"'"&_
                  ",Email='"&Email&"' where Store_Id = "&Store_ID

              end if
          else
             sql_insert = "Update Sys_Billing set Amount="&grand_total&_
                 ",Term='"&term&"',First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
                 "Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
                 "Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
                 "Card_Number='"&encrypt(cc_num)&"',Exp_Month='"&mm&"',Exp_Year='"&yy&"' "&_
                 ",Email='"&Email&"',Payment_Term="&Term&",Sys_Created='"&Now()&"', reseller_portion="&Reseller_total&" where Store_Id = "&Store_ID

              if sPaymentDescription = "Manual Payment" then
              
                sql_insert = "Update Sys_Billing set "&_
                 "First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
                 "Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
                 "Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
                 "Card_Number='"&encrypt(cc_num)&"',Exp_Month='"&mm&"',Exp_Year='"&yy&"' "&_
                 ",Email='"&Email&"' where Store_Id = "&Store_ID

              end if
          end if

     end if
     session("sql") = sql_insert
     conn_store.Execute sql_insert

     if level=1 then
        no_ecommerce=1
     else
        no_ecommerce=0
     end if
     if sPaymentDescription <> "Manual Payment" then

     sql_update = "update store_settings set Trial_Version=0,Service_Type="&level&",no_ecommerce="&no_ecommerce&" where Store_id ="&Store_id
     conn_store.Execute sql_update
     end if
     
     '*********************************************************************************************************
		  'code here to make an entry for the customer in the tbl_reseller_master
			service = term_name&" "&service	
			
			on error resume next
			sql_reseller = "Put_Reseller_Customer '"&SiteResellerID&"','"&Store_Id&"','"&term&"','"&Reseller_today&"','"&term_name&"','"&now()&"','"&service&"'"
			session("sql") = sql_reseller
                  conn_store.Execute sql_reseller

		 '*********************************************************************************************************

     Send_Mail Email,sReport_email,"Store Paid","Store " & Store_Id & " has paid " & total & " for easystorecreator "& service & " service to be activated."

    elseif Billing_Type = "Update" then
      if payment_method="eCheck" then
        sql_insert = "Update Sys_Billing set First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
                 "Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
                 "Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
                 "Bank_Name='"&BankName&"',Bank_ABA='"&encrypt(BankABA)&"',Bank_Account='"&encrypt(BankAccount)&"', "&_
                 "Acct_Type='"&acct_type&"',Org_Type='"&org_type&"',Check_Num='"&CheckSerial&"', "&_
                 "License_Num='"&DrvNumber&"',License_State='"&DrvState&"',License_Exp_Day='"&dobd&"', "&_
                 "License_Exp_Month='"&dobm&"',License_Exp_Year='"&doby&"'"&_
                  ",Email='"&Email&"', reseller_portion="&Reseller_total&" where Store_Id = "&Store_ID
      else
        sql_insert = "Update Sys_Billing set First_Name='"&first_name&"',Last_Name='"&last_name&"',"&_
              "Address='"&address&"', City='"&city&"', State='"&state&"', Zip='"&Zip&"', "&_
              "Country='"&country&"', Phone='"&Phone&"',Payment_Method='"&payment_method&"', "&_
              "Card_Number='"&encrypt(cc_num)&"',Exp_Month='"&mm&"',Exp_Year='"&yy&"' "&_
              ",Email='"&Email&"', reseller_portion="&Reseller_total&" where Store_Id = "&Store_ID
      end if
      session("sql") = sql_insert
      conn_store.Execute sql_insert

    else
	  '*********************************************************************************************************
		  'code here to make an entry for the customer in the tbl_reseller_master
			service = term_name&" "&service	
			
			on error goto 0
			sql_reseller = "Put_Reseller_Customer '"&SiteResellerID&"','"&Store_Id&"','"&term&"','"&Reseller_today&"','"&term_name&"','"&now()&"','"&service&"'"
			session("sql") = sql_reseller
                        conn_store.Execute sql_reseller

		 '*********************************************************************************************************
			
		
			
      Send_Mail email,sReport_email,"Store Paid","Store " & Store_Id & " has paid " & total & " for easystorecreator custom service." & Custom_Description
    end if

    sql_update = "Update store_settings set Custom_Amount=0,Overdue_Payment=0 where Store_Id="&Store_Id
    conn_store.Execute sql_update
    
    if sql_insert_payments<>"" then
       conn_store.Execute sql_insert_payments
    end if

    response.redirect "old_billing_info.asp?Accepted=Yes&Billing_Type="&Billing_Type
end if
%>
