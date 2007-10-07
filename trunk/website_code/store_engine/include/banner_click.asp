<!--#include file="header_noview.asp"-->

<%

if fn_get_querystring("BID")<>"" then
	BID = fn_get_querystring("BID")
	if not isNumeric(BID) then
		fn_redirect Switch_Name&"error.asp?message_id=1"
	end if
	sel_sel = "select url from store_banners WITH (NOLOCK) where store_id="&store_id&" and bann_id="&BID

	rs_store.open sel_sel, conn_store, 1, 1
	if not rs_store.eof then
		theURL = rs_store("URL")
		sql_add = "insert into store_banners_click_through (store_ID, Banner_ID, client_IP) values ("&store_id&", "&BID&", '"&Request.ServerVariables("REMOTE_ADDR")&"')"
		conn_store.execute sql_add
		rs_store.close
		fn_redirect theURL
	else
		rs_store.close
	end if
end if

%>
