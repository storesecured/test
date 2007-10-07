<%

Function deformatNVP ( nvpstr )
On Error resume next
	Dim AndSplitedArray,EqualtoSplitedArray,Index1,Index2,NextIndex
	Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
	AndSplitedArray = Split(nvpstr, "&", -1, 1)
	NextIndex=0
	For Index1 = 0 To UBound(AndSplitedArray)
	    EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
		For Index2 = 0 To UBound(EqualtoSplitedArray)
			NextIndex=Index2+1
			NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
			Index2=Index2+1
		Next
	Next
	Set deformatNVP = NvpCollection
	If Err.Number <> 0 Then
	Response.Redirect "error.asp?Message_id=101&message_add="&server.URLEncode(err.Description)
	else
	SESSION("ErrorMessage")	= Null
	End If
End Function

Function URLDecode(str)
On Error Resume Next
	 str = Replace(str, "+", " ")
        For i = 1 To Len(str)
            sT = Mid(str, i, 1)
            If sT = "%" Then
                If i+2 < Len(str) Then
                    sR = sR & _
                        Chr(CLng("&H" & Mid(str, i+1, 2)))
                    i = i+2
                End If
            Else
                sR = sR & sT
            End If
        Next

        URLDecode = sR
        	If Err.Number <> 0 Then
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLDecode")
	Response.Redirect "APIError.asp"
	else
	SESSION("ErrorMessage")	= Null
	End If
End Function
Sub ProcessOrder()

' Create an empty order
        Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
        res=order.setPartName("order")
	' Create an empty part
        Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")                
        
      ' Build 'orderoptions'
        ' For a test, set result to GOOD, DECLINE, or DUPLICATE
        'res=op.put("result", "GOOD")
        If Auth_Capture then
        	v_type = "SALE"
        else
        	v_type = "PREAUTH"
        end if
        if sType="Capture" then
		v_type="POSTAUTH"
	elseif sType = "Credit" then
		v_type="CREDIT"
	else
		v_type="VOID"
	end if
        res=op.put("ordertype", v_type)
        ' add 'orderoptions to order
        res=order.addPart("orderoptions", op)


        ' Build 'merchantinfo'
        res=op.clear()
        res=op.put("configfile", C_Configfile)
        ' add 'merchantinfo to order
        res=order.addPart("merchantinfo", op)


        ' Build 'creditcard'
        res=op.clear()

        res=op.put("cardnumber", decrypt(CardNumber))
        sArray = split(CardExpiration,"/")

        res=op.put("cardexpmonth", sArray(0))
        res=op.put("cardexpyear", sArray(1))
        ' add 'creditcard to order
        res=order.addPart("creditcard", op)

        'response.redirect Switch_Name&"error.asp?Message_id=98&Message_Add="&decrypt(CardNumber)

        ' Build 'payment'
        res=op.clear()
        res=op.put("chargetotal", GGrand_Total)
        ' add 'payment to order
        res=order.addPart("payment", op)

        ' Add oid
        res=op.clear()
        res=op.put("oid", Verified_Ref)
        ' add 'transactiondetails to order
        res=order.addPart("transactiondetails", op)


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

        ' Call LPTxn
        'resp = LPTxn.send(keyfile, host, port, outXml)
        'response.redirect Switch_Name&"error.asp?Message_id=98&Message_Add="&C_Keyfile
        resp = LPTxn.send(C_Keyfile, C_Host, C_Port, outXml)
        
        'Store transaction data on Session and redirect
        Session("outXml") = outXml
        Session("resp") = resp

        Set LPTxn = Nothing
        Set order = Nothing
        Set op    = Nothing
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
  if sApproved = "APPROVED" then
     Session("Verified_Ref") = R_Ref
     Session("AuthNumber") = R_Code
     'Session("avs_result") = R_AVS
     Session("avs_result") = ""
  else
      sError = R_Error
      sMessage = R_Message
      response.redirect Switch_Name&"error.asp?Message_id=98&Message_Add="&_
      Server.UrlEncode(R_Error&"<BR>"&R_Message)
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
