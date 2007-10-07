<%
'******************************************************************************
' Copyright (C) 2006 Google Inc.
'  
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'  
'      http://www.apache.org/licenses/LICENSE-2.0
'  
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'******************************************************************************
%>
<!--#include virtual="common/crypt.asp"-->
<!-- #include file="googlecart.asp" -->
<!-- #include file="googletax.asp" -->
<!-- #include file="googleshipping.asp" -->
<!-- #include file="googleorder.asp" -->
<!-- #include file="googlenotification.asp" -->
<!-- #include file="googlemerchantcalculation.asp" -->
<!-- #include file="xmlbuilder.asp" -->

<%
' ***IMPORTANT***
' You must define these Const variables before running the code.

' ***PLEASE READ THIS***
' Is this for Sandbox or Production environment?
' Sandbox ("SANDBOX") and Checkout ("PRODUCTION") environments are separate and 
'     require separate merchant accounts.
' Your Sandbox Merchant ID and Key cannot be used in the "PRODUCTION" environment, and
' Your Checkout production Merchant ID and Key cannot be used in the "SANDBOX" environment.
' To sign in or sign up for a Sandbox merchant account, go to 
'     https://sandbox.google.com/checkout/sell
' To sign in or sign up for a Production merchant account, go to 
'     https://checkout.google.com/sell
'
' "SANDBOX" OR "PRODUCTION"
Dim EnvType


Set rs_store_merchant = Server.CreateObject("ADODB.Recordset")
sql_real_time = "exec wsp_real_time_property "&Store_Id&",38;"
session("sql")=sql_real_time
rs_store_merchant.open sql_real_time,conn_store,1,1
rs_store_merchant.MoveFirst
Do While Not rs_store_merchant.EOF
   select case rs_store_merchant("Property")
    case "Merchant_ID"
    M_ID = trim(decrypt(rs_store_merchant("Value")))
       case "Merchant_Key"
	    Mkey = trim(decrypt(rs_store_merchant("Value")))
    case "Merchant_Currency"
       Mcurrency = trim(decrypt(rs_store_merchant("Value")))

   end select
rs_store_merchant.MoveNext
Loop
rs_store_merchant.Close
set rs_store_merchant = nothing



' Your Merchant ID and Key
' Can be found in Settings->Integration in your Merchant Center


 if M_ID <> "" then
MerchantId = cstr(M_ID)
end if
if Mkey <> "" then
MerchantKey = cstr(Mkey)
end if
if Mcurrency <> "" then
MerchantCurrency = cstr(Mcurrency)
end if

'bassel_gado@yahoo.com sandbox
'Const MerchantId = "262206148179982"
'Const MerchantKey = "6TDqzfbJ5cW9l8f6zuKjLQ"

' Currency
'Const MerchantCurrency = "USD"

 if store_id=139 or store_id=33019 or store_id=4991 then
EnvType = "SANDBOX"
else
EnvType = "PRODUCTION"
end if


currentpath = "http://" & Request.ServerVariables("SERVER_NAME")
' Cart Processing Page where the cart XML will be genearted and posted to Google
CartProcessingPage=currentpath & "/include/google/CartProcessing.asp"

' File to log Google Checkout messages
' Make sure the file permission is properly set for writing.
'Const LogFilename = "d:\logs\googlemessage.log"
'Const LogFilename = "c:\sites\store_engine\include\google\googlemessage.log"

' Google Checkout Schema URI
Const SchemaUri = "http://checkout.google.com/schema/2"

' Define PostUrl and DiagnoseUrl
Dim ServerUrl, BaseUrl, PostUrl, DiagnoseUrl

EnvType = UCase(EnvType)
If EnvType = "SANDBOX" Then
	ServerUrl = "https://sandbox.google.com/checkout/"
ElseIf EnvType = "PRODUCTION" Then
	ServerUrl = "https://checkout.google.com/"
End If
BaseUrl = ServerUrl & "cws/v2/Merchant/" & MerchantId
'response.write BaseUrl
'response.end
PostUrl = BaseUrl & "/request"
DiagnoseUrl = BaseUrl & "/request/diagnose"


'******************************************************************************
' The SendRequest function sends the request XML to the POST URL and returns
'     the response.
'
' Input:      XmlData    XML API request
'             PostUrl    URL address to which the request will be sent
' Returns:    XML response from the Google server as text
'******************************************************************************
Function SendRequest(XmlData, PostUrl)
    if M_ID="" or Mkey="" or Mcurrency="" then
    fn_error "Your Google Checkout merchant data are incomplete , please fill these data first"
    else
    LogMessage XmlData

	Dim XmlHttp, BasicAuthentication, ResponseXml
    Set XmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    XmlHttp.Open "POST", PostUrl, false

    ' Do NOT ignore Server SSL Cert Errors
    Const SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2
    Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
    XmlHttp.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, _
        (XmlHttp.getOption(SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS) - _
        SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)

    BasicAuthentication = Base64Encode(MerchantId & ":" & MerchantKey)

    XmlHttp.SetRequestHeader "Authorization", "Basic " & BasicAuthentication
    XmlHttp.SetRequestHeader "Content-Type", "application/xml; charset=UTF-8"
    XmlHttp.SetRequestHeader "Accept", "application/xml; charset=UTF-8"
    XmlHttp.Send XmlData
	ResponseXml = XmlHttp.ResponseText

	LogMessage ResponseXml

	SendRequest = ResponseXml

	Set XmlHttp = Nothing
	end if
End Function


'******************************************************************************
' The LogMessage function logs a message to a file. It also logs the time that
'     the message is logged.
'
' Input:    Message    The message to be logged
'******************************************************************************
Sub LogMessage(Message)
if 1=0 then
	Const IO_MODE = 8 ' Append
	Dim oFs, oTextFile
    Set oFs = Server.createobject("Scripting.FileSystemObject")
    Set oTextFile = oFs.OpenTextFile(LogFilename, IO_MODE, True)
    oTextFile.WriteLine now
    oTextFile.WriteLine Message
    oTextFile.Close
    Set oFS = Nothing
    Set oTextFile = Nothing
    end if
End Sub

%>

<script language="javascript" type="text/javascript" runat="server">

/** 
 * The Base64Encode function converts a string to a base64 encoded string
 *
 * Input:    input    string
 * 
 * Returns the base64 encoded string
 */
function Base64Encode(input) {

    var base64Key = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    var base64Pad = "=";
    var result = "";
    var length = input.length;
    var i = 1;

    for (i = 0; i < (length - 2); i += 3) {
        result += base64Key.charAt(input.charCodeAt(i) >> 2);
        result += base64Key.charAt(((input.charCodeAt(i) & 0x03) << 4) + (input.charCodeAt(i+1) >> 4));
        result += base64Key.charAt(((input.charCodeAt(i+1) & 0x0f) << 2) + (input.charCodeAt(i+2)>> 6));
        result += base64Key.charAt(input.charCodeAt(i+2) & 0x3f);
    }

    if (length%3) {
        i = length - (length%3);
        result += base64Key.charAt(input.charCodeAt(i) >> 2);
        if ((length%3) == 2) {
            result += base64Key.charAt(((input.charCodeAt(i) & 0x03) << 4) + (input.charCodeAt(i+1) >> 4));
            result += base64Key.charAt((input.charCodeAt(i+1) & 0x0f) << 2);
            result += base64Pad;
        } else {
            result += base64Key.charAt((input.charCodeAt(i) & 0x03) << 4);
            result += base64Pad + base64Pad;
        }
    }

    return result;
}
</script>
