<%
if Request.querystring("shopper_id")<>"" then
    Shopper_ID = Request.querystring("shopper_id")
    Session("Shopper_ID")=Shopper_ID
end if
%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
    on error goto 0
    Grand_Total = request.form("amount")
    Auth_Code = Request.Form("security_key")
    Trans_ID = Request.Form("transaction_id")
    to_email = Request.Form("to_email")
    OrderID = Request.Form("ordernumber")
    
	Verified=1
	sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",22,"&verified&",'"&Auth_Code&"','"&Trans_ID&"','','',0;"
	session("sql")=sql_update
	conn_store.execute(sql_update)
	fn_updateThirdParty=1
end function

%>
