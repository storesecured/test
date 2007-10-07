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
    str = Request.Form

    if str="" then
	    fn_error "Your payment has failed or been cancelled."
    end if

    Status= Request.Form("Status")

    if request.form("Status") = "-2" then
        fn_error "Your payment has failed or been cancelled."
    end if
  
    pay_to_email= Request.Form("pay_to_email")
    Grand_Total = request.form("mb_amount")
    Trans = split(Request.Form("transaction_id"),"-")

    OrderID = Trans(0)
    AuthNumber = Request.Form("mb_transaction_id")
    Trans_ID = Request.Form("mb_transaction_id")
    Store_Id = Request.Form("Store_Id")

    Verified=1
    sql_update = "exec wsp_purchase_verify "&Store_id&","&OrderID&",20,"&verified&",'"&AuthNumber&"','"&Trans_ID&"','','',0;"
    session("sql")=sql_update
	conn_store.execute(sql_update)
	fn_updateThirdParty=1
end function
%>
