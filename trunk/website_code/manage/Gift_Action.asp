<!--#include file="Global_Settings.asp"-->
<% 
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><%
else

	' CREATE A NEW GIFT CERTIFICATE

	if Request.Form("Form_Name") = "Create_Gift" then
		Gift_Amount = Request.Form("Gift_Amount")
		Gift_Validity = Request.Form("Gift_Validity")
		Gift_Item = Request.Form("Gift_Item")
		if not isNumeric(Gift_Amount) or not isNumeric(Gift_Validity) or not isNumeric(Gift_Item) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		Gift_name = checkStringForQ(Request.Form("Gift_Name"))
		'CHECK IF NAME IS NOT ALLREADY USED
		sql_check="select Gift_ID from store_gift_certificates where gift_name = '"&Gift_name&"' and store_id = "&store_id
		rs_store.open sql_check,conn_store
		if rs_store.bof = false then
			rs_store.close
			response.redirect "admin_error.asp?message_id=51" 
		end if
		rs_store.close
		 'RETRIEVE FORM DATA

		Gift_Restricted = checkStringForQ(Request.Form("Gift_Restricted"))
		'CHECK IF GIFT CERTIFICATE IS NOT USING AN ITEM ALLREADY
		'ASSIGNED TO ANOTHER GIFT CERTIFICATE
		sql_check="select Gift_ID from store_gift_certificates where item = '"&Gift_Item&"' and store_id = "&store_id
		rs_store.open sql_check,conn_store
		if rs_store.bof = false then
			rs_store.close
			response.redirect "admin_error.asp?message_id=54"
		end if
		rs_store.close
		 'CHECK THE AMOUT
		if cdbl(Gift_Amount)<=0 then
			response.redirect "admin_error.asp?message_id=52"
		end if
		 'CHECK VALIDITY TIME
		if cdbl(Gift_Validity)<=0 then
			response.redirect "admin_error.asp?message_id=53"
		end if
		 'INSERT GIFT CERTIFICATE INTO THE DATABASE
		sql_insert = "insert into store_gift_certificates (Store_id, Gift_Name, Amount, Validity_Time, Item, Restricted_Items) values ("&Store_id&", '"&Gift_name&"', "&Gift_Amount&", "&Gift_Validity&", '"&Gift_Item&"', '"&Gift_Restricted&"')"
		conn_store.Execute sql_insert
		sql_update = "update store_settings set Enable_Coupon = -1 where Store_id = "&Store_id
		conn_store.Execute sql_update	
		response.redirect "Gift_Manager.asp"
	end if 

	' UPDATE A GIFT CERTIFICATE
	if Request.Form("Form_Name") = "Update_Gift" then
		Gift_Amount = Request.Form("Gift_Amount")
		Gift_Validity = Request.Form("Gift_Validity")
		Gift_Item = Request.Form("Gift_Item")
		if not isNumeric(Gift_Amount) or not isNumeric(Gift_Validity) or not isNumeric(Gift_Item) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if

		'CHECK IF NAME IS NOT ALLREADY USED
		Gift_name = checkStringForQ(Request.Form("Gift_Name"))
		sql_check="select Gift_ID from store_gift_certificates where gift_id<>"&request.form("Gift_ID")&" and gift_name = '"&Gift_name&"' and store_id = "&store_id
		rs_store.open sql_check,conn_store
		if rs_store.bof = false then
			rs_store.close
			response.redirect "admin_error.asp?message_id=51" 
		end if
		rs_store.close
		'RETRIEVE FORM DATA

		Gift_Restricted = checkStringForQ(Request.Form("Gift_Restricted"))
		'CHECK IF GIFT CERTIFICATE IS NOT USING AN ITEM ALLREADY
		'ASSIGNED TO ANOTHER GIFT CERTIFICATE
		sql_check="select Gift_ID from store_gift_certificates where gift_id<>"&request.form("Gift_ID")&" and item = '"&Gift_Item&"' and store_id = "&store_id
		rs_store.open sql_check,conn_store
		if rs_store.bof = false then
			rs_store.close
			response.redirect "admin_error.asp?message_id=54" 
		end if
		rs_store.close
		'CHECK THE AMOUT
		if cdbl(Gift_Amount)<=0 then
			response.redirect "admin_error.asp?message_id=52"
		end if
		'CHECK VALIDITY TIME
		if cdbl(Gift_Validity)<=0 then
			response.redirect "admin_error.asp?message_id=53"
		end if
		'UPDATE THE GIFT CERTIFICATE
		sql_insert = "update store_gift_certificates set Gift_Name='"&Gift_name&"',  Amount="&Gift_Amount&", Validity_Time="&Gift_Validity&", Item='"&Gift_Item&"', Restricted_Items= '"&Gift_Restricted&"' where store_id="&store_id&" and gift_id="&request.form("Gift_ID")
		conn_store.Execute sql_insert
		response.redirect "Gift_Manager.asp"
	end if

	if request.queryString("Delete_Gift")<>"" then
		if not isNumeric(request.queryString("Delete_Gift")) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if

		'DELETE GIFT CERTIFICATE RECORD FROM STORE_GIFT_CERTIFICATES TABLE
		sql_del = "delete from store_gift_certificates where store_id="&store_id&" and gift_id="&request.queryString("Delete_Gift")
		conn_store.Execute sql_del
		'ALSO DELETE RELATED RECORDS FROM GIFT_CERTIFICATES_DETS TABLE
		sql_del = "delete from store_gift_certificates_dets where store_id="&store_id&" and gift_id="&request.queryString("Delete_Gift")
		conn_store.Execute sql_del
		response.redirect "Gift_Manager.asp"
	end if

End If 

%>
