<!--#include file="header_noview.asp"-->

<%
oid=fn_get_querystring("oid")

if isNumeric(oid) and fn_get_querystring("Return_from")="Admin" then
    'CHECK IF TO SEND ORDER NOTIFICATIONS EMAILS TO MERCHANT, CUSTOMER
    sql_update = "exec wsp_purchase_manual_process "&store_id&","&oid&";"
    fn_print_debug sql_update
    session("sql")=sql_update
    conn_store.execute sql_update

    fn_redirect Switch_Name&"include/check_out_payment_action.asp?Return_From=Admin"
else
    response.write "Error"
end if

%>

