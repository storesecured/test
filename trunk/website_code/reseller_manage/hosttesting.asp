<%host = replace(lcase(request.servervariables("HTTP_HOST")),"www.","")
Response.Write "host"&host
host = split(host,".")
Response.Write "here"&host(0)

%>