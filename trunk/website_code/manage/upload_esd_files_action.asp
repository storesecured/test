<!--#include file="Global_Settings.asp"-->

<%

	Server.ScriptTimeout = 500
	Download_Folder = fn_get_sites_folder(Store_Id,"Download")
	
	UploadSizeLimit = 100000000 '(100 MB max)
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.Upload
	For each file In mySmartUpload.Files
		If not file.IsMissing Then
			if file.Size > UploadSizeLimit then
				error = error & "File size exceeds limit of "&UploadSizeLimit&"|"
			end if		  
		end if		 
	Next
	if error <> "" then 
		Error_Log = error 
		%><!--#include file="include\error_template.asp"--><%
	else	
		For each file In mySmartUpload.Files
			If not file.IsMissing Then
				file.SaveAs(Download_Folder & file.FileName)
			End If
		Next
		fn_redirect "file_list.asp?"&sAddString
	end if


%>
