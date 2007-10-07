<%
var = trim(Request.QueryString("reseller"))

Response.Write "Var="&var

host = replace(lcase(request.servervariables("HTTP_HOST")),"www.","")
Response.Write "host"&host
Response.End
%>