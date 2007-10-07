<%

'On Error Resume Next

Dim  secret_key
Dim  BP_MERCHANT,BP_MODE,BP_MISSING_URL,BP_APPROVED_URL,BP_DECLINED_URL, BP_SUBMIT_URL
Dim  BP_TTYPE,BP_AMOUNT,BP_REBILLING
Dim  BP_REB_FIRST, BP_REB_EXPR, BP_REB_CYCLES, BP_REB_AMOUNT, BP_RRNO
Dim  BP_AUTOCAP,BP_AVS_ALLOWED
Dim  BP_CC_NUM,BP_CVCCVV2,BP_CC_EXPIRES,BP_ORDER_ID,BP_NAME
Dim  BP_ADDR1,BP_ADDR2,BP_CITY,BP_STATE,BP_ZIPCODE,BP_COMMENT,BP_PHONE,BP_EMAIL
Dim  BP_TAMPER_PROOF_SEAL
Dim  BP_RESPONSE

' This is the first subroutine you should call.  Required before processing
' a transaction:
Sub bp_setup(in_mid, in_key, in_mode, in_murl, in_aurl, in_durl, in_surl)
  BP_MERCHANT = in_mid
  secret_key  = in_key
  BP_MODE     = in_mode
  BP_MISSING_URL = in_murl
  BP_APPROVED_URL = in_aurl
  BP_DECLINED_URL = in_durl
  BP_SUBMIT_URL = in_surl
End Sub 

' This sets the information required for an AUTH or a SALE:
Sub bp_set_cust_required(in_cardnum, in_cardccvv2, in_cardexpire, _
                     in_name, in_addr1, in_city, in_state, in_zip)

  BP_CC_NUM     = in_cardnum
  BP_CVCCVV2    = in_cardccvv2
  BP_CC_EXPIRES = in_cardexpire
  BP_NAME       = in_name
  BP_ADDR1      = in_addr1
  BP_CITY       = in_city
  BP_STATE      = in_state
  BP_ZIPCODE    = in_zip
End Sub

' This sets up a SALE:
Sub bp_easy_sale(in_amount) 
  BP_TTYPE  = "SALE"
  BP_AMOUNT = in_amount
End Sub

' This sets up an AUTH:
Sub bp_easy_auth(in_amount)
  BP_TTYPE  = "AUTH"
  BP_AMOUNT = in_amount
End Sub

' This sets up a REFUND:
Sub bp_easy_refund(in_rrno)
  BP_TTYPE = "REFUND"
  BP_RRNO  = in_rrno
End Sub

' This sets up a REBCANCEL (rebilling cancellation)
' Input is the RRNO of the rebilling template or any transaction
' in the rebilling sequence.
Sub bp_easy_rebcancel(in_rrno)
  BP_TTYPE = "REBCANCEL"
  BP_RRNO = in_rrno
End Sub

' This sets up a CAPTURE -- input is the RRNO of the AUTH to capture.
Sub easy_capture(in_rrno)
  BP_TTYPE = "CAPTURE"
  BP_RRNO = in_rrno
End Sub

' This adds rebilling to an AUTH or a SALE.
Sub bp_easy_add_rebill(in_amount, in_first, in_expr, in_cycles)
  BP_REB_AMOUNT = in_amount
  BP_REB_FIRST  = in_first
  BP_REB_EXPR   = in_expr
  BP_REB_CYCLES = in_cycles
End Sub

' This adds AVS proofing to an AUTH.
Sub bp_easy_add_avs_proof(in_allowed)
  BP_AVS_ALLOWED = in_allowed
End Sub

' This adds Automatic Capture to an AUTH.
Sub bp_easy_add_autocap()
  BP_AUTOCAP = "1"
End Sub

' The next several functions set optional parameters.
Sub bp_set_orderid(in_oid)
  BP_ORDER_ID = in_oid
End Sub

Sub bp_set_addr2(in_addr2)
  BP_ADDR2 = in_addr2
End Sub

Sub bp_set_comment(in_com)
  BP_COMMENT = in_com
End Sub

Sub bp_set_phone(in_phone)
  BP_PHONE = in_phone
End Sub

Sub bp_set_email(in_email)
  BP_EMAIL = in_email
End Sub
  
' This function calculates your TAMPER_PROOF_SEAL for the transaction.
' It MUST BE CALLED before processing any transaction on our 1.5 interface
' or greater.
' It is called automatically by bp_process, though, so you shouldn't need to
' worry about it.
Sub bp_calc_tps()
  Dim md5Input, md5er
  Set md5er = server.createobject("FusionHashing.MD5")
  md5input = secret_key & BP_MERCHANT & BP_TTYPE & BP_AMOUNT & BP_REBILLING & _
             BP_REB_FIRST & BP_REB_EXPR & BP_REB_CYCLES & BP_REB_AMOUNT & _
             BP_RRNO & BP_AVS_ALLOWED & BP_AUTOCAP & BP_MODE
  BP_TAMPER_PROOF_SEAL = md5er.Hash(md5input)
  Set md5er = Nothing
End Sub


' This is the function that does the work of submitting to Bluepay and reading the response.
Sub bp_process()
  Dim bpreq
  Dim serverquery

  bp_calc_tps()

  Set bpreq = server.CreateObject("WinHttp.WinHttpRequest.5.1")

  ' URI Escape is Server.URLEncode(var)
  '
  serverquery = "MERCHANT="      & Server.URLEncode(BP_MERCHANT) & _
  		"&MISSING_URL="       & Server.URLEncode(BP_MISSING_URL) & _
		"&APPROVED_URL="      & Server.URLEncode(BP_APPROVED_URL) & _
		"&DECLINED_URL="      & Server.URLEncode(BP_DECLINED_URL) & _
                "&MODE="              & Server.URLEncode(BP_MODE) & _
	        "&TAMPER_PROOF_SEAL=" & Server.URLEncode(BP_TAMPER_PROOF_SEAL) & _
		"&TRANSACTION_TYPE="  & Server.URLEncode(BP_TTYPE) & _
		"&CC_NUM="            & Server.URLEncode(BP_CC_NUM) & _
		"&CVCCVV2="            & Server.URLEncode(BP_CVCCVV2) & _
		"&CC_EXPIRES="        & Server.URLEncode(BP_CC_EXPIRES) & _
		"&AMOUNT="            & Server.URLEncode(BP_AMOUNT) & _
		"&Order_ID="          & Server.URLEncode(BP_ORDER_ID) & _
		"&NAME="              & Server.URLEncode(BP_NAME) & _
		"&Addr1="             & Server.URLEncode(BP_ADDR1) & _
		"&Addr2="             & Server.URLEncode(BP_ADDR2) & _
		"&CITY="              & Server.URLEncode(BP_CITY) & _
		"&STATE="             & Server.URLEncode(BP_STATE) & _
		"&ZIPCODE="           & Server.URLEncode(BP_ZIPCODE) & _
		"&COMMENT="           & Server.URLEncode(BP_COMMENT) & _
		"&PHONE="             & Server.URLEncode(BP_PHONE) & _
		"&EMAIL="             & Server.URLEncode(BP_EMAIL) & _
		"&REBILLING="         & Server.URLEncode(BP_REBILLING) & _
		"&REB_FIRST_DATE="    & Server.URLEncode(BP_REB_FIRST) & _
		"&REB_EXPR="          & Server.URLEncode(BP_REB_EXPR) & _
		"&REB_CYCLES="        & Server.URLEncode(BP_REB_CYCLES) & _
		"&REB_AMOUNT="        & Server.URLEncode(BP_REB_AMOUNT) & _
		"&RRNO="              & Server.URLEncode(BP_RRNO) & _
		"&AUTOCAP="           & Server.URLEncode(BP_AUTOCAP) & _
		"&AVS_ALLOWED="       & Server.URLEncode(BP_AVS_ALLOWED)

' here we perform a POST; the string we've just created goes in the BODY of the POST:
  bpreq.Open "POST" ,BP_SUBMIT_URL, False
  bpreq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  bpreq.Send(serverquery)

' We return Bluepay's Response; if your program doesn't want to parse that, then
' you may use the convenience functions which follow:
  BP_RESPONSE = bpreq.GetResponseHeader("location")
  'response.write(bpreq.GetAllResponseHeaders())
  Set bpreq = Nothing
End Sub

' Returns the status: "APPROVED", "DECLINED", "MISSING", "ERROR"
Function bp_get_status()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "Result=(\d+)"
 Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_status = ExpMatched.Value
	Next
  Else
  	bp_get_status = Null
  End If
  Set ExpReg = Nothing
End Function
  
' Returns the RRNO, if any.
Function bp_get_RRNO()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "RRNO=(\d+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_RRNO = ExpMatched.Value
	Next
  Else
  	bp_get_RRNO = Null
  End If
  Set ExpReg = Nothing
End Function

' Returns the AUTHCODE, if any.
Function bp_get_Auth()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "AUTHCODE=(\d+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_Auth = ExpMatched.Value
	Next
  Else
  	bp_get_Auth = Null
  End If
  Set ExpReg = Nothing
End Function

' Returns the AVS, if any.
Function bp_get_Avs()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "AVS=(\d+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_Avs = ExpMatched.Value
	Next
  Else
  	bp_get_Avs = Null
  End If
  Set ExpReg = Nothing
End Function

' Returns the CVV2, if any.
Function bp_get_CVV2()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "CVV2=(\d+)"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_CVV2 = ExpMatched.Value
	Next
  Else
  	bp_get_CVV2 = Null
  End If
  Set ExpReg = Nothing
End Function

' Returns the message - describes the transaction.
Function bp_get_message()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern="MESSAGE=(.*?)[\&$]"
  Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_message = ExpMatched.Value
	Next
  Else
  	bp_get_message = Null
  End If
  Set ExpReg = Nothing
End Function

' Returns the name of a missing field, if the status was MISSING.
Function bp_get_missing()
  Dim ExpReg
  Set ExpReg = new RegExp
  ExpReg.pattern = "MISSING=(\w+)"
 Set ExpMatch = ExpReg.Execute(BP_RESPONSE)
  If ExpMatch.count > 0 Then
  	For each ExpMatched in ExpMatch
  		bp_get_missing = ExpMatched.Value
	Next
  Else
  	bp_get_missing = Null
  End If
  Set ExpReg = Nothing
End Function
%>
