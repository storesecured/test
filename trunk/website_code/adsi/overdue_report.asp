<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->


<%

on error resume next
sql_select = "select distinct store_id,site_name,service_type,store_email,store_phone,overdue_payment,last_login,sys_created from store_settings where overdue_payment in (5,25,55,95)"
	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

        sTextTotal="Store Id"&chr(9)&"Site Name"&chr(9)&"Overdue Days"&chr(9)&"Service Type"&chr(9)&"Email"&chr(9)&"Phone"&chr(9)&"Last Login"&chr(9)&"Sys Created"

	if noRecords = 0 then
        FOR rowcounter= 0 TO myfields("rowcount")
                sStore_id = mydata(myfields("store_id"),rowcounter)
                site_name = mydata(myfields("site_name"),rowcounter)
                service_type = mydata(myfields("service_type"),rowcounter)
                sEmail = mydata(myfields("store_email"),rowcounter)
                sPhone = mydata(myfields("store_phone"),rowcounter)
                last_login = mydata(myfields("last_login"),rowcounter)
                sys_created = mydata(myfields("sys_created"),rowcounter)
                overdue_payment = mydata(myfields("overdue_payment"),rowcounter)
                my_store_id=sStore_id

                sTextTotal=sTextTotal&vbcrlf&sStore_id&chr(9)&site_name&chr(9)&overdue_payment&chr(9)&service_type&chr(9)&sEmail&chr(9)&sPhone&chr(9)&last_login&chr(9)&sys_created

        Next
	Else
	       sTextTotal="None for today"
        End If
	

Send_Mail sReport_email,sFollow_email,"Overdue Report",sTextTotal

%>
