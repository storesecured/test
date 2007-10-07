<!--#include file="include/header.asp"-->
<!--#include file="include/sub.asp"-->
<%

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Site_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else


	 ' RETRIEVING FORM DATA
	 Invoice = checkStringForQ(request.form("Invoice"))
	 Last_Name = checkStringForQ(request.form("Last_Name"))
	 Notes = checkStringForQ(request.form("Notes"))

	 if Invoice="" then
	 	fn_error "Invoice number is required."
	 end if
	 if Last_Name="" then
	 	fn_error "Last name is required."
	 end if
	 sql_select = "select * from store_purchases WITH (NOLOCK) where Store_Id="&Store_Id&" and OID="&Invoice&" and ShipLastName='"&Last_Name&"'"
	 session("sql")=sql_select
	 rs_Store.open sql_select,conn_store,1,1
	 if rs_store.eof and rs_store.bof then
		 response.write "No matching order was found in the database please try again."
	 else
		 daysSince = datediff("d",date(),rs_Store("Sys_Created"))
		 if daysSince > Auto_RMA_Days then
			  response.write "Since it has been more than "&days&" days since your order was placed you cannot get an RMA number online automatically.  An email has been sent to the store admin requesting your RMA number."
			  Call Send_Mail(rs_Store("ShipEmail"),Store_Email,"RMA Request","Customer has requested RMA # for Order #"&Invoice & vbcrlf & Notes)

		 else
		    Randomize
			  sRmaNum = Invoice & "--"&Round((100000 * Rnd))
			  Return_Date = Now()
			  Cid=rs_Store("cid")
			  sql_insert = "Insert into Store_Purchases_Returns (Store_id,oid,cid,Return_RMA,Return_notes) values ("&Store_id&","&Invoice&","&cid&",'"&sRmaNum&"','"&Notes&"')"
			  Conn_store.execute sql_insert
			  Call Send_Mail(Store_Email,rs_Store("ShipEmail"),"RMA Request","Your RMA # is "&sRmaNum)
			  Call Send_Mail(rs_Store("ShipEmail"),Store_Email,"RMA Request","Customer has received an RMA # for Order #"&Invoice & vbcrlf & Notes)
        response.write "Your RMA number is "&sRmaNum&".	Please include this when sending back your items.	This number has been sent to your email address as well."
		 end if
	 end if

	 rs_Store.close
end if
%>
<!--#include file="include/footer.asp"-->

