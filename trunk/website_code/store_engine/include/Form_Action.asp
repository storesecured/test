<!--#include file="header_noview.asp"-->

<% 

if Request.form("Shopper_Id_Retrieve") <> "" then
	Shopper_Id = Request.form("Shopper_Id_Retrieve")
elseif fn_get_querystring("Shopper_Id_Retrieve") <> "" then
	Shopper_Id = fn_get_querystring("Shopper_Id_Retrieve")
end if

	Shopper_id_Retrieve = checkStringForQ(Shopper_Id)
	
	if not(isNumeric(Shopper_id_Retrieve)) then
		fn_error "The shopper id must be numeric."
	end if
	sql_check = "select shopper_id from store_shoppingcart WITH (NOLOCK) where Store_id = "&Store_id&" and shopper_id = "&Shopper_id_Retrieve&" and cart_processed=0"
	fn_print_debug sql_check
	rs_Store.open sql_check,conn_store,1,1
	if rs_Store.Bof = False then
		'FORWARD TO VIEW CART WITH NEW SHOPPER ID
		fn_redirect Switch_Name&"Show_Big_Cart.asp?Shopper_Id="&Shopper_id_Retrieve
	else
		'FAILURE. FORWART TO VIEW CART WITH OLD ID
		fn_redirect Switch_Name&"error.asp?message_id=29"
	end if


%>
