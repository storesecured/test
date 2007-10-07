<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="error"

%>
<!--#include file="header.asp"-->

<% 
if request.form("Update") <> "" and Request.Form("message") <> "" then
	 fromname = Replace(Request.Form("fromname"), "'", "''")
	 fromemail = Replace(Request.Form("fromemail"), "'", "''")
	 message = Replace(Request.Form("message"), "'", "''") & vbcrlf & request.form("errorinfo")

	 Set Mail = Server.CreateObject("Persits.MailSender")
	 Mail.From = fromemail
	 Mail.AddAddress "support@"&host
	 Mail.Subject = "error report form from " & fromname
	 Mail.Body = message
	 On Error Resume Next
	 Mail.Send
	 response.write "Thankyou for taking the time to help us improve."
else
	on error resume next
	sArray = split(request.querystring, ";")
	sErrorNum = sArray(0)
	sURL = sArray(1)
	sArray = split(sErrorNum,"-")
	iError = sArray(0)

	Select Case iError
		case 400
			sError = "Bad request"
		case 401
			sError = "You are not authorized to view this page"
		case 403
			sError = "You are not authorized to view this page"
		case 404
			sError = "The page cannot be found"
		case 405
			sError = "Method not allowed"
		case 406
			sError = "The resource cannot be displayed"
		case 407 
			sError = "Proxy Authentication Required"
		case 410
			sError = "The page does not exist"
		case 412
			sError = "Precondition Failed"
		case 414
			sError = "Request-URI to long"
		case 500
			sError = "Internal Server Error"
		case 501 
			sError = "Not Implemented"
		case 503
			sError = "The page cannot be displayed"
	End Select
	response.write "<B>Error</B><BR><BR>"&sError & "<BR>Please try using the provided navigation links to find what you are looking for.<BR>"&sURL

	%>
	<HR><BR><B>Help Us Improve</b><BR>
	Use the form below to tell us exactly what you were trying to do before this error occured.
	<form method="POST" action="error.asp" name="Error">
			<table>
			<tr><td>Name</td>
			<td><input type=text name=fromname size=20></td></tr>
			<tr><td>Email</td>
			<td><input type=text name=fromemail size=20></td></tr>
			<tr><td>To</td>
			<td>support@<%= host %>
			<input type=hidden name="errorinfo" value="<%= request.querystring %>"></td></tr>
	
			<tr><td colspan=2><textarea rows="5" name="message" cols="50"></textarea></td></tr>
			<tr><td colspan=2><input type="submit" name=Update value="Send"></td></tr>
			</table>
	</form>
<% end if %>
<!--#include file="footer.asp"-->
