<!--#include file = "include/header.asp"-->

<%
on error goto 0
intResellerID = session("ResellerID") 
	' Instantiate Upload Class
	DestinationPath="d:\esaleswizard\www\reseller_sites_new\logos\"
	
	Store_Id=Session("Store_Id")
    sOnlyFiles = "safe"
    
    Server.ScriptTimeout = 750
	UploadSizeLimit = 500000
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
			'if  Instr(file.ContentType,"image") = 0 then
				'Upload=0
				'error = error & "File must be an image " & file.Ext &"|"
			'End if
			filespec_a = DestinationPath&file.FileName

			if Upload=1 then
				i = i + 1
				file.SaveAs(DestinationPath & intResellerID & "_"&file.FileName)
				sFileName = intResellerID&"_"&replace(file.FileName," ","%20")
			end if
		End If
	Next

	set mySmartUpload = Nothing
	
	'code here to check whether uploading for firstime or logo is there
	
	sql_select = "select fld_title_image from tbl_reseller_logo where fld_reseller_id="&intResellerID
	set rs_store=conn.execute(sql_select)
	if not rs_store.eof and not rs_store.bof then
	    imagename=rs_store("fld_title_image")
	end if    
	rs_store.close
	
				
		if (imagename <> "") then 
			sqlUpdate = " UPDATE tbl_reseller_logo "&_
						" SET fld_logo_title='"&strlogotitle&"',fld_title_image = '"&sFileName&"'"&_
						" WHERE fld_reseller_id =" &intResellerID 
			response.Write sqlUpdate			
			conn.execute(sqlUpdate)
			intflag="2"
			
		' When Uploading a New File Previous File Should be Deleted
			'set FSO=server.CreateObject("Scripting.FileSystemObject")
			'strhidimagePath = DestinationPath &strhidimage 
			
			'if FSO.fileexists(strhidimagePath) then
			'	FSO.DeleteFile (strhidimagePath)
			'end if

			'set FSO = nothing			
			
		
	else
		sqlPutData = "insert into tbl_Reseller_Logo(fld_reseller_id,fld_logo_title,fld_title_image) VALUES" &_
				" ("&intResellerID&",'"&strlogotitle&"','"&sFileName&"')"
		response.Write sqlPutData
		conn.execute (sqlPutData)
		intflag="1"
		
	end if

response.Redirect "reseller_change_logo.asp"

%>

