<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%

sFormAction = "uncancel_store.asp"
sFormName = "payment"
sTitle = "Uncancel Store"
thisRedirect = "uncancel_store.asp"
sMenu = "account"
createHead thisRedirect 


	sql_update ="update store_settings set store_cancel=Null where store_id="&Store_Id
	conn_store.Execute sql_update

        response.write "<TR bgcolor='#FFFFFF'><TD>Your store will no longer be removed.  Thankyou for continuing your service.  We are happy to have you back.</td></tr>"

createFoot thisRedirect,0 %>

