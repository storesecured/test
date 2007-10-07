<%
if Request.Form("M_Shopper_id")<>"" then
    Shopper_ID = Request.Form("M_Shopper_id")
    Session("Shopper_ID")=Shopper_ID
end if
iDontRedirect=1
%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

iDontRedirect=1
function fn_updateThirdParty()
    on error goto 0
    transStatus = request.form("transStatus")
    if transStatus = "Y" then
        Grand_Total = request.form("authAmount")
        OrderID = Request.Form("cartId")
        Auth_Code = Request.Form("transId")
        Trans_ID = Request.Form("transId")
        'Store_Id = Request.form("M_Store_Id")
        Shopper_Id = Request.form("M_Shopper_Id")
        instId = request.form("instId")
        avs_result = request.form("AVS")
        Trans_ID = request.form("transTime")
        sCard = request.form("cardType")
        Card_Name = checkStringForQ(request.form("name"))

        Verified=1
        sql_update = "exec wsp_purchase_verify "&Store_id&","&OrderID&",17,"&verified&",'"&Auth_Code&"','"&Trans_ID&"','','',0;"
        session("sql")=sql_update
		conn_store.execute(sql_update)
    else
        fn_error "An error has occurred while processing your transaction."
    end if
    fn_updateThirdParty=1
    
	response.write "<a href='"&Switch_Name&"recipiet.asp?Shopper_Id="&Shopper_Id&"'>Click here to view the receipt</a>"

end function
%>