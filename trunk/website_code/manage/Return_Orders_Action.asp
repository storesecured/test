<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->

<%
'RETRIEVE FORM DATA
Return_RMA = Replace(Request.Form("Return_RMA"), "'", "''")
Return_notes = Replace(Request.Form("Return_notes"), "'", "''")
Return_Date = Now()
Cid = Request.Form("Cid")
Oid = Request.Form("Oid")

sql_insert = "Insert into Store_Purchases_Returns (Store_id,oid,cid,Return_RMA,Return_notes,Return_Date) values ("&Store_id&","&oid&","&cid&",'"&Return_RMA&"','"&Return_notes&"','"&Return_Date&"')"
session("sql")=sql_insert
Conn_store.execute sql_insert

'UPDATE STORE_PURCHASES TABLE WITH THE GENERAL STATUS OF SHIPPING
sql_update = "update store_purchases set Returned = 1 where Oid = "&Oid&" and Store_id= "&Store_id
Conn_store.execute sql_update

'on error resume next
if Request.Form("send_return_email") = 1 then
	' send_email to customer regards to that return :
	Return_Email_To = Request.Form("Return_Email_To")
	Return_Email_From = Request.Form("Return_Email_From")
	Return_Email_Subject = Request.Form("Return_Email_Subject")
	Return_Email_Body = Request.Form("Return_Email_Body")
	Call Send_Mail(Return_Email_From,Return_Email_To,Return_Email_Subject,Return_Email_Body)
end if

response.write Request.Form("send_return_email")
Response.Redirect "order_details.asp?oid="&oid&"&cid="&cid

%>
