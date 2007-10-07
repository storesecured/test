<%
if Request("USER2")<>"" then
    Shopper_ID = Request("USER2")
    Session("Shopper_ID")=Shopper_ID
end if
%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
    on error goto 0
    if Request("TransID")<>"" and (Request("Auth")<>"" and Request("Auth") <> "Declined") then

        OrderID = request("RefNo")
        AuthNumber = Request("Auth")
        avs = Request("AVSCode")
        card_code_verif = Request("card_code_verif")
        Trans_ID = request("TransID")
        
        if card_code_verif = "Y" then
            card_verif = "M"
        elseif card_code_verif = "N" then
            card_verif = "N"
        end if

        Verified=1
        sql_update = "exec wsp_purchase_verify "&Store_id&","&OrderID&",24,"&verified&",'"&AuthNumber&"','"&Trans_ID&"','"&card_verif&"','"&avs_result&"',0;"
        session("sql")=sql_update
		conn_store.execute(sql_update)
    else
	    fn_error "Your transaction was not processed successfully!"&request("NOTES"))&"<BR><A href=payment.asp>Click here to go back and try again</a>"

    end if
	fn_updateThirdParty=1
end function
%>


