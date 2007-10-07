<%
'Database connection file for Easystorecreator

dim conn,strconn

session.Timeout=900


set conn = server.CreateObject("adodb.connection")
strconn = "DRIVER=SQL Server;SERVER=10.235.158.138,1433;UID=melanie;PWD=tom237;DATABASE=wizard"	

conn.open strconn
on error resume next
if err.number<>0 then
	Response.Write err.description
end if



'REPLACE 'SINGLE QUOTES
function fixquotes(str)
fixquotes = replace(str&"","'","''")
end function

function checkEncode(encodeVal)

	if encodeVal<>"" then
	checkEncode = server.HTMLEncode(encodeVal)
	else
	checkEncode = encodeVal
	end if
	
end function
%>
