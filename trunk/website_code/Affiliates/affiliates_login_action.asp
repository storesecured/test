<!--#include virtual="common/connection.asp"-->

<%

'RETRIEVE AFFILIATES CODE / PASSWORD / STORE
AFFILIATE_CODE = Request.Form("AFFILIATE_CODE")
AFFILIATE_PASSWORD = Request.Form("AFFILIATE_PASSWORD")
if Request.Form("AFFILIATE_STORE")<>"" then
	AFFILIATE_STORE = Request.Form("AFFILIATE_STORE")
else
	AFFILIATE_STORE = -1
end if

'CHECK CODE / PASSWORD / STORE
sql_login = "select store_settings.store_id,store_settings.store_currency,store_affiliates.affiliate_id,store_settings.affiliate_type,store_settings.affiliate_amount from store_affiliates inner join store_settings on store_settings.store_id=store_affiliates.Store_id where Code = '"&AFFILIATE_CODE&"' and Password = '"&AFFILIATE_PASSWORD&"' and Store_settings.store_id="&AFFILIATE_STORE
rs_Store.open sql_login,conn_store,1,1
if rs_store.eof then 
	rs_Store.Close
	Session("AFFILIATE_SESSION")="FALSE"
	Response.Redirect "error.asp?Message_id=7"
end if

'LOG IN AFFILIATE
Session("AFFILIATE_CURRENCY")=rs_store("store_currency")

Session("AFFILIATE_SESSION")="TRUE"
Session("AFFILIATE_STORE_ID")=rs_store("store_id")
Session("AFFILIATE_AFFILIATE_ID")=rs_store("Affiliate_id")
Session("AFFILIATE_LOGIN_TIME")=now()
Session("Affiliate_amount") = rs_store("Affiliate_amount")
Session("Affiliate_type") = rs_store("Affiliate_type")
rs_Store.Close

response.redirect "affiliates_reports.asp"

%>
