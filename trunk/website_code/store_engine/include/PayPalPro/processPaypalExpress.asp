<%
sub create_Form_Paypal()
%>

<% 


	rs_store.open "Select oid, Grand_Total from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1 

	if not rs_store.eof then
		amount=   formatnumber(rs_store("grand_total"),2)
		orderNo = rs_store("oid")
	end if
	rs_store.close
%>	
<form action="/include/paypalpro/processPaypalExpress_action.asp" method=post>

			<input type=hidden name="token" value=<%=session("token")%>>
			<input type=hidden  name="payerId" value=<%=session("payerId")%>>
			<input type=hidden name="total_amount" value=<%=amount%>>
			<input type=hidden name="order" value=<%=orderNo%>>
			
			<input type=hidden  name="Auth_Capture" value=<%=Auth_Capture%>>
<%
end sub
%>
