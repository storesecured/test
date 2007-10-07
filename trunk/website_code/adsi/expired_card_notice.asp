<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/crypt.asp"-->
<%
on error goto 0
Server.ScriptTimeout = 4800

sql_select = "SELECT * from expired_cards"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		sStore_Id = mydata(myfields("store_id"),rowcounter)
		store_email = mydata(myfields("store_email"),rowcounter)
		email = mydata(myfields("email"),rowcounter)
		exp_month = trim(mydata(myfields("exp_month"),rowcounter))
		exp_year = trim(mydata(myfields("exp_year"),rowcounter))
		next_billing_date = mydata(myfields("next_billing_date"),rowcounter)
		payment_method = trim(mydata(myfields("payment_method"),rowcounter))
		site_name = mydata(myfields("site_name"),rowcounter)
		card_number = right(decrypt(mydata(myfields("card_number"),rowcounter)),4)

                sSubject="StoreSecured Expired Credit Card Notice"
                sText = "Dear Store Owner"&vbcrlf&vbcrlf&_
                        "This notice relates to the following account."&vbcrlf&_
                        "Site #: "&sStore_Id&vbcrlf&_
                        "Site Name: http://"&site_name&vbcrlf&vbcrlf&_
                        "Your "&Payment_Method&" card ending in "&card_number&_
                        " has the expiration date of "&exp_month&"/"&exp_year&_
                        " which has now passed.  Please update your payment information on file "&_
                        " at your earliest convenience with a credit card that is not expired."&vbcrlf&vbcrlf&_
                        "To update your payment information please login at http://manage.storesecured.com"&vbcrlf&_
                        "Select My Account-->Payments-->Update Payment Method"&vbcrlf&vbcrlf&_
                        "Your next billing date is "&next_billing_date&", please update your payment method before this time to ensure your payment can be processed."&vbcrlf&vbcrlf&_
                        "This is an automatic notice sent via the StoreSecured system."
                        
                Call Send_Mail(sNoReply_email,email,sSubject,sText)
                if store_email<>email then
                        Call Send_Mail(sNoReply_email,store_email,sSubject,sText)
                end if
                
	Next
end if

Call Send_Mail(sNoReply_email,sFollow_email,"StoreSecured Followup",sText)
response.write "done"

set myfields = Nothing



%>

