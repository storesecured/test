<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Graphics Mode Change"
sMenu = "none"
createHead thisRedirect


response.cookies("EASYSTORECREATOR")("hide_help") = request.querystring("Mode")
response.cookies("EASYSTORECREATOR").expires = DateAdd("d",60,Now())
Session("hide_help") = request.querystring("Mode")

Response.redirect Request.ServerVariables("HTTP_REFERER")
%>



<% createFoot thisRedirect, 0%>
