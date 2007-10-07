<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%
on error goto 0
Server.ScriptTimeout = 4800

sql_select = "SELECT * from new_signups"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

sText="sys_created"&chr(9)&"sStore_id"&chr(9)&"store_email"&chr(9)&"store_phone"&chr(9)&"site_name"&chr(9)&"store_domain"&chr(9)&"service_type"&chr(9)&"trial_version"&chr(9)&"store_cancel"&chr(9)&"last_login"&vbcrlf

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sys_created = mydata(myfields("sys_created"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
		store_email = mydata(myfields("store_email"),rowcounter)
		store_phone = mydata(myfields("store_phone"),rowcounter)
		site_name = mydata(myfields("site_name"),rowcounter)
		store_domain = mydata(myfields("store_domain"),rowcounter)
		service_type = mydata(myfields("service_type"),rowcounter)
		trial_version = mydata(myfields("trial_version"),rowcounter)
		store_cancel = mydata(myfields("store_cancel"),rowcounter)
		last_login = mydata(myfields("last_login"),rowcounter)

                sText=sText&sys_created&chr(9)&sStore_id&chr(9)&store_email&chr(9)&store_phone&chr(9)&site_name&chr(9)&store_domain&chr(9)&service_type&chr(9)&trial_version&chr(9)&store_cancel&chr(9)&last_login&vbcrlf
	Next
end if

sql_select = "SELECT * from waiting_to_cancel"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 and 1=0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sys_created = mydata(myfields("sys_created"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
		store_email = mydata(myfields("store_email"),rowcounter)
		store_phone = mydata(myfields("store_phone"),rowcounter)
		site_name = mydata(myfields("site_name"),rowcounter)
		store_domain = mydata(myfields("store_domain"),rowcounter)
		service_type = mydata(myfields("service_type"),rowcounter)
		trial_version = mydata(myfields("trial_version"),rowcounter)
		store_cancel = mydata(myfields("store_cancel"),rowcounter)
		last_login = mydata(myfields("last_login"),rowcounter)

         sText=sText&sys_created&chr(9)&sStore_id&chr(9)&store_email&chr(9)&store_phone&chr(9)&site_name&chr(9)&store_domain&chr(9)&service_type&chr(9)&trial_version&chr(9)&store_cancel&chr(9)&last_login&vbcrlf
	Next
end if

Call Send_Mail(sNoReply_email,sFollow_email,"StoreSecured Followup",sText)
response.write "done"

set myfields = Nothing



%>

