<!--#include file="Global_Settings.asp"-->

<%
on error resume next

sOpenPath = replace(replace(request.querystring("Open"),"#","\"),"..\","")
sOpenURL = request.querystring("Open")


	Server.ScriptTimeout = 900
	UploadSizeLimit = 500000
	
	Image_Folder = fn_get_sites_folder(Store_Id,"Images")
	Destination_Folder = Image_Folder&sOpenPath&"\" 'Folder to store

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.AllowedFilesList = "txt,swf,doc,xls,gif,jpg,jpeg,pdf,bmp,css,htm,html,xls,zip,tiff,tif,mpg,mpeg,wav,mp3,csv,wmv,wma,png,mid,eps,js,cab,rm,ram,xml"
	mySmartUpload.DeniedFilesList = "exe,bat,asp,com,dll,vbs,reg,pcd,pif,scr,msi,msp,pif,wsc,wsf,wsh,cmd,inc"
	mySmartUpload.Upload

	i = 0
	For each file In mySmartUpload.Files
		If not file.IsMissing Then
			Upload=1
			filespec_a = Destination_Folder&file.FileName

			if Upload=1 then
				i = i + 1
				file.SaveAs(Destination_Folder & file.FileName)
				
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
		Error_Log = Err.description
	%> <!--#include virtual="common/Error_Template.asp"--><%

	else
		fn_redirect "left_image_picker.asp?Open="&server.urlencode(sOpenURL)&"&"&sAddString
	end if



%>
