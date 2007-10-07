<%
if Request.Form("USER1")<>"" then
    Shopper_ID = Request.Form("USER1")
    Session("Shopper_ID")=Shopper_ID
end if
%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
    on error goto 0
    ' read post from Verisign system

    Result = Request.Form("RESULT")
    Verified_Ref = Request.Form("PNREF")
    card_code_verif = Request.Form("CSCMATCH")
    avs = Request.Form("AVSDATA")
    trans_type = Request.Form("TYPE")
    OrderID = Request.Form("INVOICE")
    
    Store_ID=Request.Form("USER2")
    AuthNumber = Request.form("AUTHCODE")
    Post = Request.Form
    

    if card_code_verif = "Y" then
	    card_verif = "M"
    elseif card_code_verif = "N" then
	    card_verif = "N"
    elseif card_code_verif = "X" then
	    card_verif = "X"
    elseif card_code_verif = "" or not Use_CVV2 then
	    card_verif = "P"
    end if

    avsaddr = Left(avs,1)
    avszip = Mid(avs,2,1)

    if avsaddr = "N" and avszip = "N" then
	    avs_result = "N"
    elseif avsaddr = "Y" and avszip = "Y" then
	    avs_result = "Y"
    elseif avsaddr = "N" and avszip = "Y" then
	    avs_result = "Z"
    elseif avsaddr = "Y" and avszip = "N" then
	    avs_result = "A"
    elseif iavs = "Y" then
	    avs_result = "G"
    elseif iavs = "X" or avszip = "X" or avsaddr = "X" then
	    avs_result = "S"
    end if

    If trans_type = "S" then
	    stype = 1
    else
	    stype = 0
    end if

    if Result = 0 and Store_Id<>"" and OrderID<>"" then
        Verified=1
        sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",8,"&verified&",'"&AuthNumber&"','"&Verified_Ref&"','"&card_verif&"','"&avs_result&"',"&stype&";"
        session("sql")=sql_update
	   conn_store.execute(sql_update)
    else
        fn_error "An error has occurred."
    end if
	fn_updateThirdParty=1
end function
%>


