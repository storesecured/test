<%
'*****************************************************************************
' Copyright 2003 LinkPoint International, Inc. All Rights Reserved.
' 
' This software is the proprietary information of LinkPoint International, Inc.  
' Use is subject to license terms.
'
'******************************************************************************    
sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"   
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
  Do While Not Rs_Store.EOF
		select case Rs_store("Property")
	  case "storename"
		  storename = decrypt(Rs_store("Value"))
		end select
		Rs_store.MoveNext
  Loop
Rs_store.Close

C_Configfile = storename
Key_Folder = fn_get_sites_folder(Store_Id,"Key") 
C_Keyfile	 = Key_Folder&"cert.pem"
Const C_Host  = "secure.linkpt.net"
Const C_Port  = 1129

if Tax_Exempt then
	TaxExempt="Y"
else
	TaxExempt="N"
end if

country_code=fn_country_code(country)
shipcountry_code=fn_country_code(ShipCountry)

Dim IsPostBack

on error resume next

 ' process order
  Call ProcessOrder()
  ' redirect to response page

  ParseResponse(Session("resp"))

  Verified_Ref = Session("Verified_Ref")
AuthNumber = Session("AuthNumber")
avs_result = Session("avs_result")

if Auth_Capture then
	trans_type = 1
else
	trans_type = 0
end if
Sub ProcessOrder()

      ' Create an empty order
        Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
        If Auth_Capture then
        	v_type = "SALE"
        else
        	v_type = "PREAUTH"
        end if
        order.setPartName("order")
	      ' Create an empty part
        Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")                

        ' Build 'orderoptions'
        if Payment_Method="eCheck" then
           res=op.put("ordertype", "SALE")
        else
            res=op.put("ordertype", v_type)
        end if
        ' set transaction result
        res=op.put("result", "LIVE")
        ' add 'orderoptions to order
        res=order.addPart("orderoptions", op)

        if Payment_Method<>"eCheck" then
           ' Build 'transactiondetails'
           res=op.clear()
           res=op.put("transactionorigin", "ECI")
           ' add 'transactiondetails to order
           res=order.addPart("transactiondetails", op)
        end if

        ' Build 'merchantinfo'
        res=op.clear()
        res=op.put("configfile", C_Configfile)
        ' add 'merchantinfo to order
        res=order.addPart("merchantinfo", op)

        ' Build 'payment'
        res=op.clear()
        Sub_Total = GGrand_Total - Shipping_Method_Price - Tax
        if Payment_Method<>"eCheck" then
          res=op.put("subtotal", Sub_Total)
          res=op.put("tax", Tax)
          res=op.put("shipping", Shipping_Method_Price)
        end if
        res=op.put("chargetotal", GGrand_Total)
        ' add 'payment to order
        res=order.addPart("payment", op)

        ' Build 'creditcard'
        if Payment_Method="eCheck" then
           BankABA = Request.Form("BankABA")
  	   BankAccount = Request.Form("BankAccount")
  	   BankName = Request.Form("BankName")
  	   CheckSerial = Request.Form("CheckSerial")
           acct_type = Request.Form("acct_type")
           org_type = Request.Form("org_type")
           if acct_type = "CHECKING" then
              acct_type = "c"
           else
               acct_type = "s"
           end if
           if org_type = "I" then
              acct_type = "p" & acct_type
           else
              acct_type = "b" & acct_type
           end if
  	   DrvState = Request.Form("DrvState")
  	   DrvNumber = Request.Form("DrvNumber")
           res=op.clear()
           res=op.put("routing", BankABA)
           res=op.put("account", BankAccount)
           res=op.put("checknumber", CheckSerial)
           res=op.put("accounttype", acct_type)
           res=op.put("dl", DrvNumber)
           res=op.put("dlstate", DrvState)
           res=op.put("bankname", BankName)
           res=op.put("bankstate", DrvState)
           ' add 'telecheck' to order
           res=order.addPart("telecheck", op)
        else
          res=op.clear()
          res=op.put("cardnumber", CardNumber)
          res=op.put("cardexpmonth", Request.Form("mm"))
          res=op.put("cardexpyear", Request.Form("yy"))
          if Use_CVV2 then
            res=op.put("cvmvalue", CardCode)
            if CardCode<>"" then
               res=op.put("cvmindicator", "provided")
            end if
          end if
          ' add 'creditcard to order
          res=order.addPart("creditcard", op)
        end if

        ' Build 'billing'
        res=op.clear()
        res=op.put("name", first_name & " " & last_name)
        res=op.put("company", server.urlencode(ShipCompany))
        res=op.put("address1", address1)
        res=op.put("address2", address2)
        res=op.put("city", city)
        res=op.put("state", state)
        ' Required for AVS. If not provided, 
        ' transactions will downgrade.			
        res=op.put("zip", zip)
        res=op.put("addrnum", address1)
        res=op.put("country", country_code)
        res=op.put("phone", ShipPhone)
        res=op.put("fax", ShipFax)
        res=op.put("email", ShipEmail)
        ' add 'billing to order
        res=order.addPart("billing", op)

        ' Build 'shipping'
        res=op.clear()
        res=op.put("name", ShipFirstname & " " & ShipLastname)
        res=op.put("address1", ShipAddress1)
        res=op.put("address2", ShipAddress2)
        res=op.put("city", ShipCity)
        res=op.put("state", ShipState)
        res=op.put("zip", ShipZip)
        res=op.put("country", shipcountry_code)


        
       ' create transaction object	
        Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")

        if (fLog = True) and ( logLvl > 0 ) Then
        
          Dim res1, resDesc
          
          'Next call return level of accepted logging in 'res1'
          'On error 'res1' contains negative number
          'You can check 'resDesc' to get error description
          'if any
          
          res = LPTxn.setDbgOpts(logFile,logLvl,resDesc,res1)
          
        End If
        
        ' get outgoing XML from 'order' object
        Dim outXml, resp
        
        outXml = order.toXML()
        outXml=replace(outXml,",","")
        outXml=replace(outXml,"&","")
' Call LPTxn
        resp = LPTxn.send(C_Keyfile, C_Host, C_Port, outXml)

        'Store transaction data on Session and redirect
        Session("outXml") = outXml
        Session("resp") = resp

        Set LPTxn = Nothing
        Set order = Nothing
        Set op    = Nothing
        Set items = Nothing        
        Set options= Nothing
        Set item    = Nothing
End Sub
        
Sub ParseResponse( rsp)
        R_Time = ParseTag("r_time", rsp)
        R_Ref = ParseTag("r_ref", rsp)
        R_Approved = ParseTag("r_approved", rsp)
        R_Code = ParseTag("r_code", rsp)
        R_Authresr = ParseTag("r_authresronse", rsp)
        
        R_OrderNum = ParseTag("r_ordernum", rsp)
        R_Message = ParseTag("r_message", rsp)
        R_AVS = ParseTag("r_avs", rsp)
        R_Error = ParseTag("r_error", rsp)
        R_TDate = ParseTag("r_tdate", rsp)
        R_AVS = ParseTag("r_avs", rsp)
        R_Tax = ParseTag("r_tax", rsp)
        R_Shipping = ParseTag("r_shipping", rsp)
        R_FraudCode = ParseTag("r_fraudCode", rsp)
        R_ESD = ParseTag("esd", rsp)
        
        
        'parsing score can get complicated and
        'depends on merchant store setting
        'you can get up to three different scores
        'from different providers.
        'Actual averaging algoritm is then up to the merchant
        'for this samples we'll use simple value which
        'we look up in following order
        ' 1. if available, use r_provideraverage 
        ' 2. if available, use r_providerone
        ' 3. use r_score
        ' 
        R_Score = ParseTag("r_provideraverage", rsp)
        if R_Score = "" Then
        R_Score = ParseTag("r_providerone", rsp)
        End if
        if R_Score = "" Then
	R_Score = ParseTag("r_score", rsp)
	' if scoring error occured this case retuns 'ERR'
	  If R_Score = "ERR" Then
	     R_Score = ""
	  End If
        End if
        
    Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")
  sApproved = R_Approved
  if sApproved = "APPROVED" or sApproved = "SUBMITTED" then
     Session("Verified_Ref") = R_OrderNum
     Session("AuthNumber") = R_Code
     'Session("avs_result") = R_AVS
     Session("avs_result") = ""
  else
      sError = R_Error
      sMessage = R_Message

      if left(sError,10)="SGS-020005" then
         sError=sError&"<BR><HR><BR>Store Owner, please ensure that you have entered a numeric storename and uploaded your linkpoint pem file.  Both must be entered exactly as given by Linkpoint or they will be rejected."
      end if
      fn_purchase_decline oid,"The transaction was rejected by the payment processor:"&sError&"<BR>"&R_Message
  end if


    End Sub

     Function ParseTag( tag ,  rsp ) 
        Dim sb 
        Dim idxSt, idxEnd 'As Integer
        
        rsp = rsp
        
        sb = "<" & tag & ">"
        idxSt = -1
        idxEnd = -1

        idxSt = InStr(rsp,sb)
        If 0 = idxSt Then
            ParseTag = ""
            Exit Function
        End If

        idxSt = idxSt + Len(sb)
        sb = "</" & tag & ">"
        idxEnd = InStr(idxSt, rsp,sb)

        If 0 = idxEnd Then
           ParseTag = ""
           Exit Function
        End If

        ParseTag = Mid(rsp, idxSt, (idxEnd - idxSt))

    End Function




%>
