<!-- #include file="functions.asp" -->
<%
if Request.Querystring("Shopper_id")<>"" then
    Shopper_ID = Request.Querystring("Shopper_id")
    Session("Shopper_ID")=Shopper_ID
end if
%>
<!--#include virtual="include/check_out_payment_action.asp"-->
<%

function fn_updateThirdParty()
	

on error resume next
	const BASE_64_MAP_INIT ="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	dim nl
		dim Base64EncMap(127)
		dim Base64DecMap(127)
	initcodecs

	sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
	  rs_Store.MoveFirst
	  Do While Not Rs_Store.EOF
		select case Rs_store("Property")
		case "VendorPassword"
			VendorPassword = decrypt(Rs_store("Value"))
		end select
	 Rs_store.MoveNext
	Loop
	Rs_store.Close

	crypt = Request.QueryString("Crypt")
	Decoded=SimpleXor(Base64Decode(crypt),VendorPassword)

    Status=getToken(Decoded,"Status")
    Grand_Total = getToken(Decoded,"Amount")
    TxCode = split(getToken(Decoded,"VendorTxCode"),"A")
    OrderID = TxCode(0)
    AuthNumber = getToken(Decoded,"TxAuthNo")
    AVSCV2=getToken(Decoded,"AVSCV2")
    VPSTxID=getToken(Decoded,"VPSTxId")
    Trans_ID=VPSTxID
   sql_ipn = "exec wsp_ipn_insert "&Store_id&","&oid&",'" & Decoded & "','Protx','" & Status & "';"
       conn_store.execute(sql_ipn)
    if Status = "ABORT" then
	    fn_error "You elected to cancel your online payment."
    elseif Status = "NOTAUTHED" then
	    fn_error "The VSP was unable to authorise your payment."
    elseif Status = "OK" then
	    Verified=1
	    sql_update = "exec wsp_purchase_verify "&Store_id&","&oid&",19,"&verified&",'"&AuthNumber&"','"&Trans_ID&"','','',0;"
         session("sql")=sql_update
		conn_store.execute(sql_update)
		else
                  fn_error "The VSP was unable to authorise your payment."
    end if
    fn_updateThirdParty=1
end function


  


%>

