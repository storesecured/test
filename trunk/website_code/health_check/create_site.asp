<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->

<%

server.scripttimeout=4800

sLocalIP = Request.ServerVariables ("LOCAL_ADDR")
if sLocalIP=sPortalIP then
    'don't setup the sites if this is the portal server
    response.Write "no sites setup, this is the portal server"
    response.end
end if
iLocalIP = fn_convert_ip_to_bigint(sLocalIP)
store_id=request.QueryString("store_id")

sCompleted = fn_website_create(Store_Id)


%>
