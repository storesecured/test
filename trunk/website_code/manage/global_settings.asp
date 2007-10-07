<!--#include virtual="common/connection.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<!--#include file="include\sub.asp"-->
<%

if Session("Store_Id") ="" then
    sCookies = checkCookies
    if sCookies=1 then
      sPageName = Server.URLEncode(Request.ServerVariables("script_name")&"?"&request.querystring)
      Response.Redirect "Login_store.asp?ReturnTo="&sPageName
    else
      Response.Redirect "disabled_cookies.asp"
    end if
else
    store_id=session("store_id")
%>
<!--#include virtual="common/get_settings.asp"--> 
<%
end if
%>
