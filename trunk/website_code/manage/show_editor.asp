<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Graphics Mode Change"
sMenu = "none"
createHead thisRedirect


if request.querystring("Status")="Show" then
   response.cookies("EASYSTORECREATOR")("show_editor") = 1
elseif request.querystring("Status")="Old" then
   response.cookies("EASYSTORECREATOR")("show_editor") = 2
else
   response.cookies("EASYSTORECREATOR")("show_editor") = 0
end if
response.cookies("EASYSTORECREATOR").expires = DateAdd("d",360,Now())

Response.redirect Request.ServerVariables("HTTP_REFERER")
%>



<% createFoot thisRedirect, 0%>
