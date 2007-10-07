<!--#include file="Global_Settings.asp"-->
<% if request.querystring("Filename") <> "" then %>
<SCRIPT LANGUAGE = "JavaScript">

	<!--
	function setParentResults(resArg,resVal) {
		window.parent.opener.setResults(resArg, resVal);
		window.parent.close();}
	//-->

</SCRIPT>
<body onLoad=setParentResults('<%= request.queryString("returnArg") %>','<%= request.querystring("Filename") %>');>
<%
else
  
  
  on error resume next

	Server.ScriptTimeout = 750
	UploadSizeLimit = 500000
	Destination_Folder = fn_get_sites_folder(Store_Id,"Root")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.AllowedFilesList = "txt,swf,doc,xls,gif,jpg,jpeg,pdf,bmp,css,htm,html,xls,zip,tiff,tif,mpg,mpeg,wav,mp3,csv,wmv,wma,png,js"
	mySmartUpload.DeniedFilesList = "exe,bat,asp,com,dll,vbs,reg,pcd,pif,scr,msi,msp,pif,wsc,wsf,wsh,cmd,inc"
	mySmartUpload.Upload

	i = 0
	For each file In mySmartUpload.Files
		If not file.IsMissing Then
			Upload=1
			if file.Size > UploadSizeLimit then
				Upload=0
				error = error & "File size exceeds limit of "&UploadSizeLimit&"|"
			End if
			filespec_a = Destination_Folder&file.FileName

			if Upload=1 then
				i = i + 1
				file.SaveAs(Destination_Folder & file.FileName)
				sFileName = replace(file.FileName," ","%20")
			end if
		End If
	Next

	set mySmartUpload = Nothing
	if error <> "" then
		if i > 0 then
			Error_Log = error & i & " image(s) were uploaded successfully"
		else
			Error_Log = error
		end if

	elseif Err Then
	  if Err.number = -2147220494 then
	     Error_log = "Please choose a valid file to upload."
	  elseif err.number = -2147467259 then
	     Error_log = "The file type is considered unsafe and is not allowed to be uploaded."
	  elseif err.number = -2147220489 then
	     Error_log = "The file type is considered unsafe and is not allowed to be uploaded."
	  elseif err.number = 9 then
	     Error_log = "Please choose a valid file to upload."
	  else
		    Error_Log = Err.description & err.number
		end if
 %> 
  <!--#include virtual="common/Error_Template.asp"-->
 <%

	else
		fn_redirect "file_support_upload_action.asp?returnArg="&request.querystring("returnArg")&"&Filename="&server.urlencode(sFileName)
	end if
end if
%>
