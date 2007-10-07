<!--#include file="header_noview.asp"-->
<!--#include file="emails.asp"-->
<!--#include file="Send_Invoice_By_Mail.asp"-->
<%

loid=fn_get_querystring("oid")
cid=fn_get_querystring("cid")

if cid<>"" then
            SQL_select = "exec wsp_customer_lookup "&Store_Id&","&cid&",0;"
            fn_print_debug SQL_select
            rs_store.open SQL_select,conn_store,1,1
            if not rs_store.eof and not rs_store.bof then
            user_id = rs_store("user_id")
            password = rs_store("password")
            Protected_page_access = rs_Store("Protected_page_access")
            first_name = rs_store("first_name")
            last_name = rs_store("last_name")
            company = rs_store("company")
            address1 = rs_store("address1")
            address2 = rs_store("address2")
            city = rs_store("city")
            state = rs_store("state")
            zip = rs_store("zip")
            country = rs_store("country")
            email = rs_store("email")
            phone = rs_store("phone")
            fax = rs_store("fax")
            spam = rs_store("spam")
            is_residential = rs_store("is_residential")
            ccid = rs_store("ccid")
            orders_dept = rs_store("orders_dept")
            tax_exempt = rs_store("tax_exempt")
            budget_left = rs_store("budget_left")
            reward_left = rs_store("reward_left")
            else
		  	fn_redirect "The customer record no longer exists."
		  end if
		  rs_store.close
            Groups=fn_Get_Cid_Groups(cid)
        end if

if isNumeric(loid) then
    send_order_notifications(loid)
end if
response.write "<BR>Notification(s) sent for order "&loid
response.write "<BR>Close this window to exit"

%>

