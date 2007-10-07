<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<html>

<head>

<%
server.scripttimeout=4800
on error resume next

'find all of the stores which have not been logged into in 6 weeks, are trial stores which have been sent cancellation warnings.
sql_select = "SELECT Store_Id from Store_Settings where ((DATEDIFF(DAY,Last_Login,'"&Now()&"') >= 43 and Warning_Sent=3) or (DATEDIFF(DAY,Last_Login,'"&Now()&"') >= 7 and Last_login=Sys_created)) and Trial_Version=1 and (server="&freeserver&" or server="&paidserver&")"


set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sStore_Id = mydata(myfields("store_id"),rowcounter)
                Delete_Store sStore_Id
		sText = sText & chr(13) & chr(10) & "Deleted Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
	Next
end if

'find all of the stores which the owner requested a cancellation for at least 2 days ago
sql_select = "SELECT Store_Id from Store_Settings where DATEDIFF(DAY,Store_Cancel,'"&Now()&"') >= 2 and (server="&freeserver&" or server="&paidserver&")"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sStore_Id = mydata(myfields("store_id"),rowcounter)
		Delete_Store sStore_Id
		sText = sText & chr(13) & chr(10) & "Deleted Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
	Next
end if


'find all of the stores which the owner requested a cancellation for at least 2 days ago
sql_select = "SELECT Store_Id from Store_Settings where overdue_payment=98 and warning_sent=4"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sStore_Id = mydata(myfields("store_id"),rowcounter)
		Delete_Store sStore_Id
		sText = sText & chr(13) & chr(10) & "Deleted Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
	Next
end if

function Delete_Store (Del_Store)
        sql_insert="insert into sys_adsi (adsi_command,adsi_options,adsi_server) values ('Delete_Web','"&Del_Store&"','10.235.158.133')"
        conn_store.execute sql_insert
        sql_insert="insert into sys_adsi (adsi_command,adsi_options,adsi_server) values ('Delete_Web','"&Del_Store&"','10.235.158.134')"
        conn_store.execute sql_insert
        sql_insert="insert into sys_adsi (adsi_command,adsi_options,adsi_server) values ('Delete_Dir','"&Del_Store&"','10.235.158.132')"
        conn_store.execute sql_insert
        sql_insert="insert into sys_adsi (adsi_command,adsi_options,adsi_server) values ('Delete_Site','"&Del_Store&"','10.235.158.132')"
        conn_store.execute sql_insert
end function


%>
</font>
</body>
</html>
