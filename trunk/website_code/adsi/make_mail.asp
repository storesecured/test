<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%

 'mailserver vars
MailServerIP="127.0.0.1"
AuthUserName="blac6789"
AuthPassword="easystore"
ServerIP="192.168.1.20"
				
' Create a server connection object

sql_select = "select store_id, store_domain,store_domain2,store_password,email_offsite from store_settings where mail_done=0 order by store_domain"
response.write sql_select

set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

	if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sDomain = replace(lcase(mydata(myfields("store_domain"),rowcounter)),"www.","")
		store_password = replace(mydata(myfields("store_password"),rowcounter),"www.","")
		sstore_id = replace(mydata(myfields("store_id"),rowcounter),"www.","")
		email_offsite = replace(mydata(myfields("email_offsite"),rowcounter),"www.","")

		' check if domain already exists
		
		response.write "email_offsite="&email_offsite
		response.write "sDomain="&sDomain
		if not(email_offsite) and sDomain<>"" and not isNull(sDomain) then
		      response.write "in first if"
		      if not IsDomainName(sDomain) then
			response.write "in 2nd if"
			sURL = "http://"&MailServerIP&"/services/svcDomainAdmin.asmx"
			sURL = sURL & "/AddDomain"

			sParams = "AuthUserName="&AuthUserName&"&AuthPassword="&AuthPassword
			sParams = sParams & "&DomainName="& sDomain&"&Path="&Mail_Path&sDomain&"&PrimaryDomainAdminUserName=managedomain&PrimaryDomainAdminPassword="&store_password&"&PrimaryDomainAdminFirstName=None&PrimaryDomainAdminLastName=None"                      ' string
			sParams = sParams & "&IP="&ServerIP&"&ImapPort=143&PopPort=110&SMTPPort=25&MaxAliases=0&MaxDomainSizeInMB=0&MaxDomainUsers=0&MaxMailboxSizeInMB=0&MaxMessageSize=5000&MaxRecipients=100&MaxDomainAliases=5&MaxLists=5&ShowDomainAliasMenu=True&ShowContentFilteringMenu=True&ShowSpamMenu=True&ShowStatsMenu=True&RequireSmtpAuthentication=True&ShowListMenu=True&ListCommandAddress=stServ"                         ' string

			Response.Write "Request: " & sURL & "<br>"
			response.write "sParams: "&sParams
			iStatusCode = CallWebServicePost(sURL, sParams, sReply)

			if iStatusCode <> 200 then
			  Response.Write "Add domain failed with status: " & iStatusCode & ".  <BR>URL = " & sURL & "    <BR>Response = " & sReply
			   response.write "<br> Add domain failed for "&sDomain 
			else
				   response.write "<br> Add domain successful for "&sDomain 	
				Response.Write "<br><br>Response: " & sReply & "<br><br>"
				
				sql_update="update store_settings set mail_done=1 where store_id="&sstore_id
				conn_store.execute sql_update
				set oDom = NewDomDoc(sReply)

				if oDom.documentElement.selectSingleNode("ResultCode").text <> "0" then
					Response.Write "<b>Add domain FAILED:</b> <pre>" & oDom.xml & "</pre><br>"
				else
					Response.Write "<b>Add domain SUCCEEDED:</b> <pre>" & oDom.xml & "</pre><br>"

				end if
				set oDom = nothing
			end if

		      end if
		 end If
		 sql_update="update store_settings set mail_done=1 where store_id="&sstore_id
			response.write "<BR>"&sql_update
			conn_store.execute sql_update
	next
	end if



'======================================
'for checking if domain names exists in SmarterMail
Function IsDomainName(DomainName)

	sMainURL = "http://"&MailServerIP&"/Services/svcDomainAdmin.asmx"

	sParams = "AuthUserName="&AuthUserName&"&AuthPassword="&AuthPassword
	sParams = sParams & "&DomainName=" &DomainName				' string

	sURL = sMainURL & "/GetDomainInfo"

	iStatusCode = CallWebServicePost(sURL, sParams, sReply)

	if iStatusCode <> 200 then
		Response.Write "<BR>Get Domain failed with status: " & iStatusCode& ".  URL = " & sURL & "    Response = " & sReply
		IsDomainName=false
	else
		Response.Write "Response: " & sReply & "<br>"
		set oDom = NewDomDoc(sReply)
		if oDom.documentElement.selectSingleNode("ResultCode").text <> "0" then
			Response.Write "<b>Get Domain FAILED:</b> <pre>" & oDom.xml & "</pre><br>"
			IsDomainName=false
		else
			Response.Write "<font size=2 face=arial color=#008000><b>Get  Domain SUCCEEDED:</b></font><font face=arial><pre>" & oDom.xml & "</pre></font><br>"
			IsDomainName=true
		end if
		set oDom = nothing
	end If

End Function 



'===========================================
Function AddDomainAlias(DomainName, DomainAliasName)

	'mailserver vars
	'MailServerIP="127.0.0.1:9998"
	'AuthUserName="admin"
	'AuthPassword="admin"

	sMainURL = "http://"&MailServerIP&"/Services/svcDomainAliasAdmin.asmx"

	sParams = "AuthUserName="&AuthUserName&"&AuthPassword="&AuthPassword
	sParams = sParams & "&DomainName=" &DomainName				' string
	sParams = sParams & "&DomainAliasName=" &DomainAliasName				' string
	
	sURL = sMainURL & "/AddDomainAliasWithoutMxCheck"
	
	iStatusCode = CallWebServicePost(sURL, sParams, sReply)

	if iStatusCode <> 200 then
		'Response.Write "<BR>Add Domain Alias failed with status: " & iStatusCode& ".  URL = " & sURL & "    Response = " & sReply
		AddDomainAlias=false
	else
		'Response.Write "Response: " & sReply & "<br>"
		set oDom = NewDomDoc(sReply)
		if oDom.documentElement.selectSingleNode("ResultCode").text <> "0" then
			'Response.Write "<b>Add Domain Alias FAILED:</b> <pre>" & oDom.xml & "</pre><br>"
			AddDomainAlias=false
		else
			'Response.Write "<font size=2 face=arial color=#008000><b>Add  Domain Alias SUCCEEDED:</b></font><font face=arial><pre>" & oDom.xml & "</pre></font><br>"
			AddDomainAlias=true
		end if
		set oDom = nothing
	end If


End Function 











'=============basic dom/post functions=====================
function NewDomDoc(a_sXml)
  dim oDom
  set oDom = CreateObject("MSXML2.DomDocument")
  oDom.async = false
  oDom.loadXML(a_sXml)
  set NewDomDoc = oDom
  set oDom = nothing
end function


function CallWebServicePost(a_sURL, a_sPostData, a_sReply)
  dim objXML
  set objXML = CreateObject("MSXML2.ServerXMLHTTP")  
  objXML.open "POST", a_sURL, false  
  objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  objXML.send a_sPostData
  a_sReply = objXML.responseText
  CallWebServicePost = objXML.status
  set objXML = nothing  
end function

%>
