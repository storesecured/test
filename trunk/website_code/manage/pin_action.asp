<!--#include file="Global_Settings.asp"-->
	
<%
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if
	
'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><% 
else
	BID = request.form("BID")
	Item_Sku = request.form("Item_Sku")
	pin = request.form("pin")
	other_info = checkStringForQ(request.form("other_info"))
	pin_used = request.form("pin_used")
	oid = request.form("oid")
  cid= ""

	if pin_used = "on" then
		pin_used = -1
	else
		pin_used = 0
	end if
	
	if isNumeric(oid) then
		sql_select_cid= "select cid from Store_Purchases where store_id="&store_id&" and oid='"&oid&"'"
		rs_store.open sql_select_cid, conn_store, 1, 1
	  if not rs_store.bof and not rs_store.eof then
				cid= rs_store("cid")
		end if
		rs_store.close
  end if
		
  if request.form("Action")="Add" then
	 		if pin_used = 0  then
				sql_add = "insert into store_pin (store_id,item_sku,pin,other_info,pin_used) values ("&store_id&",'"&item_sku&"', '"&pin&"', '"&other_info&"', "&pin_used&")"
			else
				sql_add = "insert into store_pin (store_id,item_sku,pin,other_info,pin_used,cid,use_date,oid) values ("&store_id&",'"&item_sku&"', '"&pin&"', '"&other_info&"', "&pin_used&", '"&cid&"', convert(char(12),getdate(),101), '"&oid&"')"
			end if
	Else
			if pin_used = 0 then
	 			sql_add = "update store_pin set item_sku='"&item_sku&"',pin='"&pin&"',other_info='"&other_info&"',pin_used="&pin_used&" where store_id="&store_id&" and itemdown_id="&BID
			else
				sql_add = "update store_pin set item_sku='"&item_sku&"',pin='"&pin&"',other_info='"&other_info&"',pin_used="&pin_used&",cid='"&cid&"',use_date=convert(char(12),getdate(),101),oid = '"&oid&"' where store_id="&store_id&" and itemdown_id="&BID
			end if
	end if
  session("sql") = sql_add
	
	conn_store.execute sql_add
	response.redirect "edit_pin1.asp"
	
end if
%>
