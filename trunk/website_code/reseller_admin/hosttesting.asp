<%host = replace(lcase(request.servervariables("HTTP_HOST")),"www.","")
Response.Write "host"&host
host = split(host,".")
if ubound(host)> 1 then 
	site = host(1)
else
			site = host(0)
end if 
Response.Write "herwe"&site



%>