<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sFormAction = "cancel_store_action.asp"
sFormName = "payment"
sTitle = "Cancelled Store"
thisRedirect = "cancel_store_action.asp"
sMenu = "account"
createHead thisRedirect 



 sql_select = "select Payment_Method from sys_billing where store_id="&Store_Id
	rs_store.open sql_select, conn_store, 1, 1
	if not rs_store.eof then
		Payment_Method = replace(rs_Store("Payment_Method")," ","")
	end if
	rs_store.close

	response.write "<TR bgcolor='FFFFFF'><TD>Your request has been received and will be processed automatically within the next 24-48 hours.</td></tr>"
	if (Payment_Method = "Paypal" or Payment_Method = "PayPal") and Service_Type > 0 then
		response.write "<TR bgcolor='FFFFFF'><TD>To completely cancel your store please cancel your PayPal subscription payment by logging into your PayPal account.</td></tr>"
	end if


createFoot thisRedirect,0 %>

