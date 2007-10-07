<!--#include virtual="common/connection.asp"-->

<!--#include file="include/sub.asp"-->
<%


sql_select = "select count(site_name) as stores,sys_affiliates.email_address from store_settings inner join sys_affiliates on store_settings.affiliate_id=sys_affiliates.affiliate_id where service_type>=7 and trial_version=0 and overdue_payment=0 and store_settings.affiliate_id>8 group by sys_affiliates.email_address order by sys_affiliates.email_address"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
FOR rowcounter= 0 TO myfields("rowcount")
   Pay_Amount = mydata(myfields("stores"),rowcounter) * 5
   Email_Address = mydata(myfields("email_address"),rowcounter)
   sText = sText & chr(13) & chr(10) & Email_Address&" - Pay $" & Pay_Amount
Next

sText = sText & chr(13) & chr(10) & "*********************************************"

sql_select = "select count(site_name) as stores,sys_affiliates.email_address from store_settings inner join sys_affiliates on store_settings.affiliate_id=sys_affiliates.affiliate_id where service_type<7 and service_type>=3 and trial_version=0 and overdue_payment=0 and store_settings.affiliate_id>8 group by sys_affiliates.email_address order by sys_affiliates.email_address"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
FOR rowcounter= 0 TO myfields("rowcount")
   Pay_Amount = mydata(myfields("stores"),rowcounter) * 2
   Email_Address = mydata(myfields("email_address"),rowcounter)
   sText = sText & chr(13) & chr(10) & Email_Address&" - Pay $" & Pay_Amount
Next
set myfields= nothing
Send_Mail sNoReply_email,sReport_email,"Affiliate Report",sText

response.write sText

%>

