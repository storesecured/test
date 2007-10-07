<%
if Request.Querystring("Shopper_id")<>"" then
    Shopper_ID = Request.Querystring("Shopper_id")
    Session("Shopper_ID")=Shopper_ID
end if
%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
	on error goto 0
    if request.form("err") <> "" then
        fn_error "An error has occured with your order.  Please contact seller to resolve." & request.form("verbage")
    end if

    Grand_Total = request.form("Amount")
    Auth_Code = Request.Form("ApprovalCode")
    Trans_ID = Request.Form("receiptnumber")
    MerchantNumber = Request.Form("MerchantNumber")
    sArray = split(request.form("Products"), "::")
    Shopper_Id=request.querystring("Shopper_Id")

	Verified=1
    sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",12,"&verified&",'"&Auth_Code&"','"&Trans_ID&"','','',0;"
    session("sql")=sql_update
	conn_store.execute(sql_update)
	fn_updateThirdParty=1
end function
%>
