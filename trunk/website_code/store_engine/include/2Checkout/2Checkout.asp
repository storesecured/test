<%
if Request.Form("Shopper_id")<>"" then
    Shopper_ID = Request.Form("Shopper_id")
    if instr(Shopper_ID,",")>0 then
        sPos = instr(Shopper_ID,",")
        Shopper_ID=left(Shopper_ID,sPos-1)
    end if
    Session("Shopper_ID")=Shopper_ID
end if

%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
	on error goto 0
    fn_print_debug "form="&request.form
    session("sql")=request.form
    Grand_Total = request.form("total")
    loid = Request.Form("cart_order_id")
    Auth_Code = Request.Form("key")
    Trans_ID = Request.Form("order_number")
    Order_Num = Request.Form("order_number")
    
    Verified=1
    sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",11,"&verified&",'"&Auth_Code&"','"&Trans_ID&"','','',0;"
    session("sql")=sql_update
	conn_store.execute(sql_update)
	fn_updateThirdParty=1
end function


%>
