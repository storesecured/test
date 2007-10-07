<!--#include file = "../include/headeroutside.asp"-->
<%
dim strfirst,strlast,strinsert,currentdate,straddress,strcity,strstate,strzip,strcountry,strphone
'Retrive the values for text boxes
currentdate = date()
'Response.Write "currentdate="&currentdate
strfirst = trim(Request.Form("First_name"))
strlast = trim(Request.Form("Last_name"))
straddress = trim(Request.Form("Address"))
strcity = trim(Request.Form("City"))
strstate=  trim(Request.Form("State"))
strzip = trim(Request.Form("Zip"))
strcountry = trim(Request.Form("Country"))
strphone = trim(Request.Form("Phone"))
strfax = trim(Request.Form("Fax"))
strcompany =trim(Request.Form("Company"))
strlogin = trim(Request.Form("Login"))
strpassword = trim(Request.Form ("Password"))
strconfirm = trim(Request.Form("password_confirm"))
stremail = trim(Request.Form("Email"))
strhost = trim(Request.Form("Site_Host"))


'Response.Write "strhost"&strhost
'Code here to insert the information of reseller into the database
strinsert = "Put_reseller_info '"&strfirst&"','"&strlast&"','"&straddress&"','"&strcity&"','"&strstate&"','"&strzip&"','"&strphone&"','"&strfax&"','"&strcompany&"','"&strlogin&"','"&strpassword&"','"&stremail&"','"&strhost&"','"&currentdate&"'"

conn.execute(strinsert)
set rsNew =  conn.execute("SELECT @@IDENTITY")
	if not rsNew.eof then
		resellerid = trim(rsNew(0))
	end if
Session("resellerid") = resellerid
Response.Redirect "reseller_payment.asp"
Response.End



							
%>	