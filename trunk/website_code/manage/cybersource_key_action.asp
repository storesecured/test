<!--#include file="Global_Settings.asp"-->

<%
on error goto 0


	Server.ScriptTimeout = 900
	UploadSizeLimit = 500000
	
	Key_Folder	 = fn_get_sites_folder(Store_Id,"Key")
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

   If not fso.FolderExists(Key_Folder) Then
		set f1 = fso.CreateFolder(Key_Folder)
	 end if

	mySmartUpload.AllowedFilesList = "p12"
	mySmartUpload.DeniedFilesList = "exe,bat,asp,com,dll,vbs,reg,pcd,pif,scr,msi,msp,pif,wsc,wsf,wsh,cmd,inc"
	mySmartUpload.Upload

	i = 0
	For each file In mySmartUpload.Files
		If not file.IsMissing Then
			Upload=1
			filespec_a = Key_Folder&file.FileName

			if Upload=1 then
				i = i + 1
                file.SaveAs(Key_Folder & file.FileName)
			end if
		End If
	Next

	set mySmartUpload = Nothing
	if error <> "" then
		if i > 0 then
			Error_Log = error & i & " Cybersource Key has been uploaded successfully"
		else
			Error_Log = error&Key_Folder & file.FileName
		end if
	elseif Err Then
		Error_Log = Err.description
	%> <!--#include virtual="common/Error_Template.asp"--><%

	else
		fn_redirect Request.ServerVariables("HTTP_REFERER")
	end if

%>
