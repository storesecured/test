<!--#include file="include/header_noview.asp"-->

<%
oid=fn_get_querystring("oid")

if isNumeric(oid) and fn_get_querystring("Return_from")="Admin" then
  'CHECK IF TO SEND ORDER NOTIFICATIONS EMAILS TO MERCHANT, CUSTOMER
  sql_select_orders = "select cid,ccid,oid,shopper_id from store_purchases WITH (NOLOCK) where (oid="&oid&") and store_id="&store_id
        rs_Store.open sql_select_orders,conn_store,1,1
	if rs_Store.Bof = False then
                cid=rs_store("cid")
                ccid=rs_store("ccid")
                rs_store.close
                
                sql_update="update store_shoppers set oid="&oid&",cid="&cid&",ccid="&ccid&" where store_id="&store_id&", shopper_id="&shopper_id&";"
                conn_store.execute sql_update
                
                sql_update = "update store_shoppingcart set cid="&cid&" where store_id="&store_id&" and shopper_id="&shopper_id&";"
                conn_store.execute sql_update
                response.redirect "check_out_action.asp?Return_From=Admin&Shopper_Id="&Shopper_Id&"&Oid="&Oid
        else
                rs_store.close
                response.write "could not find matching purchase"
        end if


else
response.write "You are not authorized to view this page"
end if
%>

