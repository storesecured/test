<!--#include file="include/header.asp"-->
<%

function fn_get_error (sQString)
	sSanitize=request.querystring(sQString)
	sSanitize=fn_replace(sSanitize,"<script","")
     sSanitize=fn_replace(sSanitize,"</script>","")

     fn_get_error=sSanitize
end function

Result = fn_get_error("Result")
Error = fn_get_error("Error")
if Error="" then
   response.write Result
else
   response.write "Error: "&Error
end if
%>
<BR><BR><a href=trackform.asp class=link>Click here to track more packages</a>




<!--#include file="include/footer.asp"-->
