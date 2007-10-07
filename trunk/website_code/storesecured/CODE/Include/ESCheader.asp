<%

'code here to check for the session of the reseller

			
if trim(session("Escid"))="" then 
%>

<script language="javascript">
	alert("Your session has expired.\n Please start from the Login page")
	document.location.href = "ESc_login.asp"
</script>
<%	
end if



'Database connection file for Easystorecreator

dim conn,strconn

session.Timeout=900


set conn = server.CreateObject("adodb.connection")
strconn = "DRIVER=SQL Server;SERVER=198.87.87.59;UID=user;PWD=thinkmore;DATABASE=wizard"	

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