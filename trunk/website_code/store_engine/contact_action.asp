<!--#include file="include/header.asp"-->
<%

'ERROR CHECKING

If not CheckReferer then
	fn_redirect Switch_Name&"error.asp?message_id=1"
end if

If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else	
	fromname = checkStringForQ(Request.Form("fromname"))
	fromemail = checkStringForQ(Request.Form("fromemail"))
	message = Request.Form("message")

	'CHECK EMAIL VALIDITY
	if fn_IsValidEmail(fromemail)=0 then
		fn_redirect Switch_Name&"error.asp?Message_Id=33"
	end if
 
	Call Send_Mail(fromemail,Store_Email,"Contact Form from " & fromname,message)
	
	response.write "Your email has been successfully sent."
End If 

%>
<!--#include file="include/footer.asp"-->
