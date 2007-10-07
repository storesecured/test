<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()

    select case left(Request.QueryString("Response"), 7)
    case "RSP0000"
        mesg="Your check has been APPROVED"
    case "RSP0001"
        mesg= "Your check has been DECLINED"
    case "RSP0010"
        mesg= "Test Completed"
    case "RSP0020"
        mesg= "This check has already been approved before. Duplicate Transaction"
    case "RSP1101"
        mesg= "Please enter a check number greater than 99"
    case "RSP1102"
        mesg= "Please enter a check amount greater than $2.00"
    case "RSP1201"
        mesg= "Please enter a name"
    case "RSP1202"
        mesg= "Please enter your address"
    case "RSP1203"
        mesg= "Please enter your city"
    case "RSP1204"
        mesg= "Please enter your state"
    case "RSP1205"
        mesg= "Please enter your zip"
    case "RSP1301"
        mesg= "Please enter the name of your bank"
    case "RSP1302"
        mesg= "Please enter the city your bank is located in"
    case "RSP1303"
        mesg= "Please enter the state your bank is located in"
    case "RSP1304"
        mesg= "Please enter the zip your bank is located in"
    case "RSP1313"
        mesg= "MICR number is either invaild or missing"
    case "RSP1401"
        mesg= "Please enter a valid ID number"
    case "RSP1402"
        mesg= "Please enter the state of issuance on your ID"
    case "RSP1501"
        mesg= "Please enter your phone number including area code"
    case "RSP1502"
        mesg= "Please enter your email address for a confirmation"
    case "RSP9999"
        mesg= "The ChecksByNet service is not available at this time. Please use a different payment method."
    case else
        mesg= "There has been an error in processing your request. Please try again. "&Request.QueryString("Response")
    end select

    Result = Request.querystring("Response")
    order_id = Request.querystring("OrderID")
    Shopper_id=request.QueryString("Shopper_id")

    Result = left(Result,7)

    if Result = "RSP0010" or Result = "RSP0000" then    
        sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",33,"&verified&",'"&AuthNumber&"','"&AuthNumber&"','','',0;"
        session("sql")=sql_update
		conn_store.execute(sql_update)
    else
	    fn_error mesg
    end if
    fn_updateThirdParty=1
end function
%>

