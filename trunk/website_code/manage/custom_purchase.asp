<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->


<%
sFormAction = "custom_purchase.asp"
sTitle = "Custom Purchase"
thisRedirect = "custom_purchase.asp"
sMenu="none"
createHead thisRedirect

sAmount=request.querystring("amount")
sDescription=request.querystring("description")

sql_update = "update store_settings set Custom_Amount="&sAmount&",custom_description='"&sDescription&"' where Store_id ="&Store_id
conn_store.Execute sql_update

response.redirect "custom.asp"

createFoot thisRedirect, 0 %>


