<!--#include file="include/header.asp"-->

<% 
If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	fn_error Error_Log
	fn_redirect Switch_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else
if request.form <> "" then
send_to = request.form("Send_To")
if send_to = "" then
	send_to = Store_Email
end if
subject = request.form("Subject")
if subject = "" then
	subject = Store_name & " Form Post"
end if
Success_Message = request.form("Success_Message")
if Success_Message = "" then
	Success_Message = "Your form has been submitted successfully."
end if

on error goto 0
sText =""
For ix = 1 to request.form.count
   Key=request.form.key(ix)     
   if (Right(Key,2) <> "_C") then
      if Key <> "Subject" and Key<>"Success_Message" and Key<>"Send_To" then
         if Right(Key,6) = "_Match" then
             sMatchKey = replace(Key,"_Match","")
             if request.form(key) <> request.form(sMatchKey) then
                fn_error "The 2 "&sMatchKey&" entries must match."
             end if
         else
             sText = sText & (Key & " = " & fn_sanitize(request.form.item(ix)) & vbcrlf)
         end if
      end if
   end if

next
sText=sText & (vbcrlf & "IP Address = " & Request.ServerVariables("REMOTE_ADDR") & vbcrlf)
sText=sText & ("Page = " & Request.ServerVariables("HTTP_REFERER") & vbcrlf)

Send_Mail Store_Email,send_to,subject,sText
end if

%>
<Table><tr><td><%= Success_Message %></td></tr></table>
<% end if %>
<!--#include file="include/footer.asp"-->
