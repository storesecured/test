<!--#include virtual="common/global_settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
sFormAction = "httperror.asp"
sTitle = "Error"
sSubmitName = "Store_Activation_Update"
thisRedirect = "httperror.asp"
sTopic="Error"
sMenu = "none"
createHead thisRedirect
%>
<TR><TD>
<%

blnShowUserTheError = True
  blnCaptureQSValuesForEmail = True
  blnShowUserTheQSValues = False
  blnCaptureFormValuesForEmail = True
  blnShowUserTheFormValues = False
  blnCaptureSessionValuesForEmail = True
  blnShowUserTheSessionValues = False
  blnCaptureApplicationValuesForEmail = True
  blnShowUserTheApplicationValues = False
  blnSendAdminMail = True
  blnAppendLog = False
  blnAppendDatabase =False

  strServerName = Request.ServerVariables("Server_Name")
  strServerPort = Request.ServerVariables("Server_Port")
  strScriptName = Request.ServerVariables("Script_Name")
  strURL = "http://" & strServerName & ":" & strServerPort & "/" & strScriptName
  strScriptLink = "<a href=" & strURL & ">" & strURL & "</a>"
  strCRLF = Chr(13) & Chr(10) 'Create the paragraph break
  strSectionDelimiter = "#########################################" & strCRLF
  strDelimiter = "|" 'delimiter for use in the error details string for
'appending to the log

Set objASPError = Server.GetLastError()

  'assign variable names to property values
  strErrorNo = CStr(objASPError.Number)
  strErrorCode = CStr(objASPError.ASPCode)
  strErrorDescription = CStr(objASPError.Description)
  strASPDescription = CStr(objASPError.ASPDescription)
  strCategory = CStr(objASPError.Category)
  strFileName = CStr(objASPError.File)
  strLineNo = CStr(objASPError.Line)
  strCol = CStr(objASPError.Column)
  strErrorSource = CStr(objASPError.Source)

  if blnShowUserTheError = True then
	ShowUserTheErrorDetails()
  end if

  if blnCaptureQSValuesForEmail = True then
	CaptureQSValuesForEmail()
  end if

  if blnShowUserTheQSValues = True then
	ShowUserTheQSValues()
  end if

  if blnCaptureFormValuesForEmail = True then
	CaptureFormValuesForEmail()
  end if

  if blnShowUserTheFormValues = True then
	ShowUserTheFormValues()
  end if

  if blnCaptureSessionValuesForEmail = True then
	CaptureSessionValuesForEmail()
  end if

  if blnShowUserTheSessionValues = True then
	ShowUserTheSessionValues()
  end if

  if blnCaptureApplicationValuesForEmail = True then
	CaptureApplicationValuesForEmail()
  end if

  if blnShowUserTheApplicationValues = True then
	ShowUserTheApplicationValues()
  end if

  if blnAppendLog = True then
	AppendErrorLog()
  end if

  if blnAppendDatabase = True then
	AppendDatabase()
  end if

  if blnSendAdminMail = True then
	ConfigurationReminder()
	SendAdminEmail()
	SendTxtEmail()
  end if

Function ShowUserTheErrorDetails()
	if strErrorNo = 0 then
	   'no error exists, nothing to report
	   response.redirect "admin_home.asp"
	
   elseif strErrorNo = -2147467259 then
	   'transaction deadlock, wait a random amount of time and try again
      Randomize
      iRnd =  1000 * Rnd
      response.write "Please wait"
      for i=0 to iRnd
         response.write ". ."
      next
      response.redirect strURL & "?" & request.querystring
   else %>
	<P><font face="Arial, Helvetica, sans serif" size="3">An Error of the type <B>"<%=strCategory%>"</b> has occurred.
  <BR><BR><B></b></font></P>
	<P><font face="Arial, Helvetica, sans serif" size="2">There is no need to contact us about this error as we are already aware.
	This is for information only.  <BR>Regards<BR>The <%=strServerName%> Web Master</font></P>
	<table border="1" bordercolorlight="#000000" bordercolordark="#ffffff" bgcolor="#cccccc" cellpadding="0" cellspacing="0">
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">Error Number</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strErrorNo <> "" then
		 Response.Write strErrorNo
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">Error Code</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strErrorCode <> "" then
		 Response.Write strErrorCode
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">Error Description</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strErrorDescription <> "" then
		 Response.Write strErrorDescription
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">ASP Description</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strASPDescription <> "" then
		 Response.Write strASPDescription
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	 <TR>
		<TD><font face="Arial, Helvetica, sans serif" size="1">Category</FONT></TD>
		<TD><font face="Arial, Helvetica, sans serif" size="1"><%
		if strCategory <> "" then
		 Response.Write strCategory
		else
		 Response.Write " "
		end if
		%></font></TD></TR>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">File Name</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strFileName <> "" then
		 Response.Write strFileName
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">Line Number</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strLineNo <> "" then
		 Response.Write strLineNo
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">Column</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strCol <> "" then
		 Response.Write strCol
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	  <tr>
		 <td><b><font face="Arial, Helvetica, sans serif" size="1">Error Source</font></b></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%
		if strErrorSource <> "" then
		 Response.Write strErrorSource
		else
		 Response.Write " "
		end if
		%></font></td>
	  </tr>
	</table>
<%	end if %>
	<%
  End Function

  Function CaptureQSValuesForEmail()
	if Request.QueryString <> "" then ' loop through all the querystring values and write them to the email
	 strQSEmailValues = strCRLF & strSectionDelimiter
	 strQSEmailValues = "QUERYSTRING VALUES (if any) FOLLOW:  " & strCRLF & strCRLF
	 For each x in Request.QueryString
		strQSEmailValues = strQSEmailValues & x & " : " & Request.QueryString(x) & strCRLF
	 next
	 strQSEmailValues = strQSEmailValues & strCRLF & strCRLF
	 strQSEmailValues = strQSEmailValues & "The full URL is: " & strURL & "?" & Request.QueryString
	end if
  End Function

  Function ShowUserTheQSValues()%>
	<P><font face="Arial, Helvetica, sans serif" size="2"><B>The querystring values (if any) are:</b></font></P>
	<table border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" cellpadding="0" cellspacing="0">
	  <tr>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Name</b></font></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Value</b></font></td>
	  </tr>
	  <%
	 if Request.QueryString <> "" then ' loop through all the querystring values and write them to the email
	  For each x in Request.QueryString%>
		<tr>
		  <td><font face="Arial, Helvetica, sans serif" size="1"><%=x%></font></td>
		  <td>
		  <%if Request.QueryString(x) <> "" then%>
		  <font face="Arial, Helvetica, sans serif" size="1"><%=Request.QueryString(x)%></font>
		  <%else%>

		  <%end if%></td>
		</tr>
	  <%
	  next
	 end if
	  %>

	</table>
	<%
  End Function

  Function CaptureFormValuesForEmail()
	if Request.Form <> "" then  ' loop through all the form values and write them to the email
	 strFormEmailValues = strCRLF & strSectionDelimiter
	 strFormEmailValues = strFormEmailValues & "FORM VALUES (if any) FOLLOW:  " & strCRLF & strCRLF
	 For each x in Request.Form
		strFormEmailValues = strFormEmailValues & x & " : " & Request.Form(x) & strCRLF
	 next
	end if
  End Function

  Function ShowUserTheFormValues()%>
	<P><font face="Arial, Helvetica, sans serif" size="2"><B>The form values (if any) are:</b></font></P>
	<table border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" cellpadding="0" cellspacing="0">
	  <tr>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Name</b></font></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Value</b></font></td>
	  </tr>
	  <%
	 if Request.Form <> "" then ' loop through all the querystring values and write them to the email
	  For each x in Request.Form%>
		<tr>
		  <td><font face="Arial, Helvetica, sans serif" size="1"><%=x%></font></td>
		  <td>
		  <%if Request.Form(x) <> "" then%>
		  <font face="Arial, Helvetica, sans serif" size="1"><%=Request.Form(x)%></font>
		  <%else%>

		  <%end if%></td>
		</tr>
	  <%
	  next
	 end if
	  %>

	</table>
	<%
  End Function

  Function CaptureSessionValuesForEmail()
	' loop through all the session variable values and write them to the email
	strSessionEmailValues = strCRLF & strSectionDelimiter
		strSessionEmailValues = strSessionEmailValues & "SESSION VARIABLE VALUES (if any) FOLLOW:  " & strCRLF & strCRLF
		For Each sessitem in Session.Contents
		  strSessionEmailValues = strSessionEmailValues & sessitem & " : " & Session.Contents(sessitem) & strCRLF
		Next
  End Function

  Function ShowUserTheSessionValues()%>
	<P><font face="Arial, Helvetica, sans serif" size="2"><B>The session variable values (if any) are:</b></font></P>
	<table border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" cellpadding="0" cellspacing="0">
	  <tr>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Name</b></font></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Value</b></font></td>
	  </tr>
	  <%
	 For each sessitem in Session.Contents%>
	  <tr>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%=sessitem%></font></td>
		 <td>
		 <%if Session.Contents(sessitem) <> "" then%>
		 <font face="Arial, Helvetica, sans serif" size="1"><%=Session.Contents(sessitem)%></font>
		 <%else%>

		 <%end if%></td>
	  </tr>
	 <%
	 next
	  %>

	</table>
	<%
  End Function

  Function CaptureApplicationValuesForEmail()
	' loop through all the session variable values and write them to the email
	strApplicationEmailValues = strCRLF & strSectionDelimiter
		strApplicationEmailValues = strApplicationEmailValues & "APPLICATION VARIABLE VALUES (if any) FOLLOW:  " & strCRLF & strCRLF
		For Each appitem in Application.Contents
		  strApplicationEmailValues = strApplicationEmailValues & appitem & " : " & Application.Contents(appitem) & strCRLF
		Next
  End Function

  Function ShowUserTheApplicationValues()%>
	<P><font face="Arial, Helvetica, sans serif" size="2"><B>The Application variable values (if any) are:</b></font></P>
	<table border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" cellpadding="0" cellspacing="0">
	  <tr>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Name</b></font></td>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><b>Value</b></font></td>
	  </tr>
	  <%
	 For each appitem in Application.Contents%>
	  <tr>
		 <td><font face="Arial, Helvetica, sans serif" size="1"><%=appitem%></font></td>
		 <td>
		 <%if Application.Contents(appitem) <> "" then%>
		 <font face="Arial, Helvetica, sans serif" size="1"><%=Application.Contents(appitem)%></font>
		 <%else%>

		 <%end if%></td>
	  </tr>
	 <%
	 next
	  %>

	</table>
	<%
  End Function


  Function ConfigurationReminder()

	'build an error handler configuration string for display at the foot of the email to remind us how we have configured it
	strErrorHandlerConfig = strCRLF & strSectionDelimiter
	strErrorHandlerConfig = strErrorHandlerConfig & "****Error Handler Configuration Details****" & strCRLF 'section title

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Show Client Error Details = "
	if blnShowUserTheError = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Show User The QueryString Values = "
	if blnShowUsertheQSValues = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Capture QueryString Values = "
	if blnCaptureQSValuesForEmail = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Show User The Form Values = "
	if blnShowUserTheFormValues = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Capture Form Values = "
	if blnCaptureFormValuesForEmail = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Show User The Session Variable Values = "
	if blnShowUserTheSessionValues = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Capture Application Variable Values = "
	if blnCaptureApplicationValuesForEmail = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Show User The Application Variable Values = "
	if blnShowUserTheApplicationValues = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Capture Session Variable Values = "
	if blnCaptureSessionValuesForEmail = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Send Administrative Email = "
	if blnSendAdminMail = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes" & strCRLF
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Append Error Log = "
	if blnAppendLog = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes, succeeded=" & blnAppendLogSuccessfully
	 if blnAppendLogSuccessfully = False then
	  strErrorHandlerConfig = strErrorHandlerConfig & ", reason: " & strErrDesc & strCRLF
	 else
	  strErrorHandlerConfig = strErrorHandlerConfig & strCRLF
	 end if
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

	strErrorHandlerConfig = strErrorHandlerConfig & "	  **Append Database = "
	if blnAppendDatabase = True then
	 strErrorHandlerConfig = strErrorHandlerConfig & "yes, succeeded=" & blnAppendDatabaseSuccessfully
	 if blnAppendLogSuccessfully = False then
	  strErrorHandlerConfig = strErrorHandlerConfig & ", reason: " & strErrDesc & strCRLF
	 else
	  strErrorHandlerConfig = strErrorHandlerConfig & strCRLF
	 end if
	else
	 strErrorHandlerConfig = strErrorHandlerConfig & "no" & strCRLF
	end if

  End Function

  Function AppendErrorLog()
	On Error Resume Next 'we need this because if we error out here the script stops
	'build error string for appending to log
	strErrorDetails = strErrorNo & strDelimiter & strErrorCode & strDelimiter & strErrorDescription & strDelimiter & strASPDescription & strDelimiter & strCategory & strDelimiter & strFileName & strDelimiter & strLineNo & strDelimiter & strCol & strDelimiter & strErrorSource
	'set up variables
	strPathOfFile = "C: emp" 'you must add the IUSR_YourMachineName account to the permissions for your chosen directory (backslash is required)"
	strNameOfFile = "scripterrorlog.txt"
	Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFileSystem.OpenTextFile(strPathOfFile & strNameOfFile, 8, True)
	if Err.number = 0 then objTextStream.WriteLine strErrorDetails
	if Err.number = 0 then
	 objTextStream.Close
	 blnAppendLogSuccessfully = True 'flag to report in admin email how things went
	else
	 blnAppendLogSuccessfully = False 'flag to report in admin email how things went
	 strErrDesc = Err.description
	 objTextStream.Close
	end if
	On Error GoTo 0 'restore default error checking
  End Function

  Function AppendDatabase()
	strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= yourdatabasepath;" 'database connectionstring
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open strConnectionString
	Set rsAddNewRow = Server.CreateObject("ADODB.Recordset")
	rsAddNewRow.Open "tbl_ErrorLogs", objConn, adOpenDynamic, adLockOptimistic, adCmdTableDirect
	With rsAddNewRow
	 .AddNew
	  .Fields("ErrorNo") = strErrorNo
	  .Fields("ErrorCode") = strErrorCode
	  .Fields("ErrorDescription") = strErrorDescription
	  .Fields("ASPDescription") = strASPDescription
	  .Fields("Category") = strCategory
	  .Fields("FileName") = strFileName
	  .Fields("LineNo") = strLineNo
	  .Fields("Col") = strCol
	  .Fields("ErrorSource") = strErrorSource
	 .Update
	End With
	rsAddNewRow.Close
	objConn.Close
	Set rsAddNewRow = Nothing
	Set objConn = Nothing
  End Function

  Function SendAdminEmail()
	'define variables
	strFromName = "Script Error Function" 'Email from name
	strFromAddress = sDeveloper_email
	strRemoteHost = "localhost"
	strSubjectText = "Error Occurred on: " & strServerName

	sErrorInfo = ""
	for each name in request.servervariables
		 sErrorInfo = sErrorInfo & vbcrlf & name & "="& request.servervariables(name)
	next

		strBodyText = "An error of type """ & strCategory & """ has occurred at the web site: " & strServerName _
			& strCRLF _
			& strCRLF _
			& "The URL of the error was : " & strURL _
			& strCRLF _
			& strCRLF _
			& "The error details are as follows:" _
			& strCRLF _
			& "Error Number		  :  " & strErrorNo _
			& strCRLF _
			& "Error Code			  :  " & strErrorCode _
			& strCRLF _
			& "Error Description   :  " & strErrorDescription _
			& strCRLF _
			& "ASP Description	  :  " & strASPDescription _
			& strCRLF _
			& "Category 			  :  " & strCategory _
			& strCRLF _
			& "File Name			  :  " & strFileName _
			& strCRLF _
			& "Line Number 		  :  " & strLineNo _
			& strCRLF _
			& "Column				  :  " & strCol _
			& strCRLF _
			& "Source				  :  " & strErrorSource _
			& strCRLF _
			& strSectionDelimiter & strCRLF _
			& strCRLF _
			& "Store ID="&Store_Id _
			& strCRLF _
			& sErrorInfo _
			& strCRLF _
			& session("sql")



		' attempt to send the mail and notify the client if it fails - if it 'doesn't then notify that the webmaster has been notified - you will need to
  'change this code if you do not use ASPMail
	 if strErrorNo <> 0 then
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.From = strFromAddress
		Mail.AddAddress strFromAddress
		Mail.Subject = strSubjectText
		Mail.Body = strBodyText & strQSEmailValues & strFormEmailValues & strSessionEmailValues & strApplicationEmailValues &strErrorHandlerConfig
		Mail.Queue=True
		On Error Resume Next
		Mail.Send
	 end if

  End Function
  
  Function SendTxtEmail()

      

  End Function
%>
</td></tr>
<% createFoot thisRedirect, 0%>
