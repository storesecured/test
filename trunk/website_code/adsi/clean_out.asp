<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%
on error goto 0
Server.ScriptTimeout = 4800

'find all of the stores which have not been logged into in 2 weeks, are trial stores and have not yet been sent cancellation warnings.
sql_select = "SELECT Store_Email,Store_User_ID, Store_Password, Site_Name,Store_Id from Store_Settings where DATEDIFF(DAY,Last_Login,'"&Now()&"') >= 14 and Warning_Sent=0 and Trial_Version=1"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sEmail = mydata(myfields("store_email"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
				sBody = "This message is to inform you that you have not logged into your free StoreSecured store for at least 2 weeks." &_
				"	Free stores which are not accessed at least once in a 6 week period will be automatically deleted."&vbcrlf&vbcrlf&"If you wish to keep your " &_
				"store located at http://" & mydata(myfields("site_name"),rowcounter) & ", please login sometime within the next 4 weeks.	This is the first " &_
				"warning before deletion." & vbcrlf & vbcrlf &_
				"Login: " &  mydata(myfields("store_user_id"),rowcounter) & vbcrlf &_
				"Password: " & mydata(myfields("store_password"),rowcounter) & vbcrlf &_
				"Login at: http://manage.storesecured.com" & vbcrlf & vbcrlf & _
				"If you have decided not to use StoreSecured or wish to remove your store immediately "&_
				" please go to http://manage.storesecured.com/cancel_store.asp"

		Call Send_Mail(sNoReply_email,sEmail,"StoreSecured Inactivity Warning",sBody)
		sText = sText & chr(13) & chr(10) & "Sent 1st cancellation warning to Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)

		sql_update = "Update Store_Settings set Warning_Sent=1 where store_Id = "&sStore_Id
		conn_store.Execute sql_update
	Next
end if

'find all of the stores which have not been logged into in 4 weeks, are trial stores which have been sent cancellation warnings.
sql_select = "SELECT Store_Email,Store_User_ID, Store_Password, Site_Name,Store_Id,Store_Phone,Store_State,Store_Country,Last_Login,Sys_created from Store_Settings where DATEDIFF(DAY,Last_Login,'"&Now()&"') >= 28 and Warning_Sent=1 and Trial_Version=1"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sEmail = mydata(myfields("store_email"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
				sBody = "This message is to inform you that you have not logged into your free StoreSecured store for at least 4 weeks." &_
				"	Free stores which are not accessed at least once in a 6 week period will be automatically deleted."&vbcrlf&vbcrlf&"If you wish to keep your " &_
				"store located at http://" & mydata(myfields("site_name"),rowcounter) & ", please login within the next 2 weeks.	This is the second " &_
				"warning before deletion.  Once your store is removed you will be unable to access it or any of the data contained within." & vbcrlf & vbcrlf &_
				"Login: " &  mydata(myfields("store_user_id"),rowcounter) & vbcrlf &_
				"Password: " & mydata(myfields("store_password"),rowcounter) & vbcrlf &_
				"Login at: http://manage.storesecured.com" & vbcrlf & vbcrlf & _
				"If you have decided not to use StoreSecured or wish to remove your store immediately "&_
				" please go to http://manage.storesecured.com/cancel_store.asp"

		sText = sText & chr(13) & chr(10) & "Sent 2nd cancellation warning to Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
		sql_update = "Update Store_Settings set Warning_Sent=2 where store_Id = "&sStore_Id
		conn_store.Execute sql_update
   next
end if


'find all of the stores which have not been logged into in 6 weeks, are trial stores which have been sent cancellation warnings.
sql_select = "SELECT Store_Email,Store_User_ID, Store_Password, Site_Name,Store_Id from Store_Settings where DATEDIFF(DAY,Last_Login,'"&Now()&"') >= 42 and Warning_Sent=2 and Trial_Version=1"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sEmail = mydata(myfields("store_email"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
				sBody = "This message is to inform you that you have not logged into your free StoreSecured store for at least 6 weeks." &_
				"	Free stores which are not accessed at least once in a 6 week period will be automatically deleted."&vbcrlf&vbcrlf&"If you wish to keep your " &_
				"store located at http://" & mydata(myfields("site_name"),rowcounter) & ", please login today.	This is the final " &_
				"warning before deletion.  Once your store is removed you will be unable to access it or any of the data contained within." & vbcrlf & vbcrlf &_
				"Login: " &  mydata(myfields("store_user_id"),rowcounter) & vbcrlf &_
				"Password: " & mydata(myfields("store_password"),rowcounter) & vbcrlf &_
				"Login at: http://manage.storesecured.com" & vbcrlf & vbcrlf & _
				"If you have decided not to use StoreSecured simply do nothing and your store will be removed tomorrow."

		Call Send_Mail(sNoReply_email,sEmail,"StoreSecured Deletion Warning",sBody)
		sText = sText & chr(13) & chr(10) & "Sent final cancellation warning to Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
		sql_update = "Update Store_Settings set Warning_Sent=3 where store_Id = "&sStore_Id
		conn_store.Execute sql_update
   next
end if

'find all of the stores which are overdue more then 30 days and have not been sent warnings.
sql_select = "SELECT Store_Email,Store_User_ID, Store_Password, Site_Name,Store_Id from Store_Settings where overdue_payment>91 and overdue_payment<100 and Warning_Sent<>4 and warning_sent<>5"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sEmail = mydata(myfields("store_email"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
				sBody = "This message is to inform you that your StoreSecured store is now over 1 month overdue." &_
				"Stores which are overdue more then 2 months will be automatically deleted."&vbcrlf&vbcrlf&"If you wish to keep your " &_
				"store located at http://" & mydata(myfields("site_name"),rowcounter) & ", please login today and submit payment.  This is the first " &_
				"warning before deletion.  Once your store is removed you will be unable to access it or any of the data contained within." & vbcrlf & vbcrlf &_
				"Login: " &  mydata(myfields("store_user_id"),rowcounter) & vbcrlf &_
				"Password: " & mydata(myfields("store_password"),rowcounter) & vbcrlf &_
				"Login at: http://manage.storesecured.com" & vbcrlf & vbcrlf & _
				"If you have decided not to use StoreSecured you may login and request that it be cancelled or just wait and do nothing."

		Call Send_Mail(sNoReply_email,sEmail,"StoreSecured Deletion Warning",sBody)
		sText = sText & chr(13) & chr(10) & "Sent overdue cancellation warning to Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
		sql_update = "Update Store_Settings set Warning_Sent=4 where store_Id = "&sStore_Id
		conn_store.Execute sql_update
   next
end if


'find all of the stores which are overdue more then 98 days and have been sent warnings.
sql_select = "SELECT Store_Email,Store_User_ID, Store_Password, Site_Name,Store_Id from Store_Settings where overdue_payment=99 and Warning_Sent=4"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sEmail = mydata(myfields("store_email"),rowcounter)
		sStore_Id = mydata(myfields("store_id"),rowcounter)
				sBody = "This message is to inform you that your StoreSecured store is now over 3 months overdue." &_
				"Stores which are overdue more then 1 month will be automatically deleted."&vbcrlf&vbcrlf&"If you wish to keep your " &_
				"store located at http://" & mydata(myfields("site_name"),rowcounter) & ", please login today and submit payment.  This is the final " &_
				"warning before deletion.  Once your store is removed you will be unable to access it or any of the data contained within." & vbcrlf & vbcrlf &_
				"Login: " &  mydata(myfields("store_user_id"),rowcounter) & vbcrlf &_
				"Password: " & mydata(myfields("store_password"),rowcounter) & vbcrlf &_
				"Login at: http://manage.storesecured.com" & vbcrlf & vbcrlf & _
				"If you have decided not to use StoreSecured simply do nothing and your store will be removed automatically."

		Call Send_Mail(sNoReply_email,sEmail,"StoreSecured Deletion Final Notice",sBody)
		sText = sText & chr(13) & chr(10) & "Sent overdue cancellation warning 2 to Store:"& sStore_Id & " URL:" & mydata(myfields("site_name"),rowcounter)
		sql_update = "Update Store_Settings set store_cancel='"&now()&"',warning_sent=5 where store_Id = "&sStore_Id
		conn_store.Execute sql_update
   next
end if


set myfields = Nothing
response.write sText


%>

