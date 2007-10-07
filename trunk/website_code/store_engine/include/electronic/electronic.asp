<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst	 
Do While Not Rs_Store.EOF 
	select case Rs_store("Property")
		case "ePNAccount"
			ePNAccount = decrypt(Rs_store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

Response.Buffer = True

'Create & initialize the xObjHTTP object
Set xObj = CreateObject("SOFTWING.ASPtear")

sRemoteURL = "https://www.eProcessingNetwork.Com/cgi-bin/tdbe/transact.pl"


'Send the request to the eProcessingNetwork Transparent Database Engine
Post_String = ""
Post_String = Post_String &"ePNAccount="&Server.UrlEncode(ePNAccount)
Post_String = Post_String &"&CardNo="&Server.UrlEncode(CardNumber)
Post_String = Post_String &"&Cvv2="&Server.UrlEncode(CardCode)
Post_String = Post_String &"&ExpMonth="&Server.UrlEncode(Request.Form("mm"))
Post_String = Post_String &"&ExpYear="&Server.UrlEncode(Request.Form("yy"))
Post_String = Post_String &"&Total="&Server.UrlEncode(GGrand_Total)
Post_String = Post_String &"&Address="&Server.UrlEncode(address1)
Post_String = Post_String &"&Zip="&Server.UrlEncode(zip)
Post_String = Post_String &"&HTML=No"
Post_String = Post_String &"&Inv=Report"

'xObj.Send Post_String
sResponse = xObj.Retrieve(sRemoteURL, 1, Post_String, "", "")

'parse the response string and handle appropriately
sApproval = mid(sResponse, 2, 1)

if sApproval = "Y" then
    sSplitResponse = split(sResponse,",")
    sAVSFull=sSplitResponse(1)
    sAvsSplit=split(sAvsFull,"(")
    avs_result=left(sAvsSplit(1),1)
    sAuthFull=sSplitResponse(0)
    sAuthSplit=split(sAuthFull," ")
    AuthNumber=replace(sAuthSplit(2),"""","")
    Verified_Ref=replace(sSplitResponse(4),"""","")
elseif sApproval = "N" then
    	fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&mid(sResponse, 3, 16)
else
    	fn_purchase_decline oid,"The transaction was rejected by the payment processor:<BR>"&sResponse
end if


Set xObj = Nothing
%>



