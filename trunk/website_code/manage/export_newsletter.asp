<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
server.scripttimeout = 400
on error goto 0
sCountItems = 0

sFormAction = "export_newsletter.asp"
sTitle = "Export Newsletter"
sName = ""
sFormName = "Newsletter"
sSubmitName = "Update"
thisRedirect = "export_newsletter.asp"
'sMenu = "importexport"
createHead thisRedirect

	sql_select = "select Newsletter_Id, Sys_Created, Store_Id, Email_Address, First_Name, Last_Name from Store_Newsletter where Store_id="&Store_id

	rs_Store.open sql_select,conn_store,1,1

	If rs_Store.eof = false then
		arrNews = rs_Store.getrows()
		rs_Store.movefirst
		TxtExporthead = "Newsletter ID" & chr(9)&_
									"Email Address" & chr(9)&_
									"First Name" & chr(9)&_
									"Last Name"
		TxtExport=TxtExporthead & vbcrlf

		For i = 0 to Ubound(arrNews,2)
			Newsletter_Id = arrNews(0,i)
			Email_Address = arrNews(3,i)
			First_Name = arrNews(4,i)
			Last_Name = arrNews(5,i)
			TxtExportDept =Newsletter_Id & chr(9) &_
			              removeLines(Email_Address) &  chr(9) &_
			              removeLines(First_Name) &  chr(9) &_
			              Last_Name

			TxtExportDept=TxtExportDept & vbcrlf
			TxtExport=TxtExport&TxtExportDept
		Next
	End If
	set myFields=nothing

Set FileObject = CreateObject("Scripting.FileSystemObject")

File_Name = "newsletters_"&dateserial(year(now()),month(now()),day(now()))
File_Name = Replace(File_Name,"/","-")
eFile_Name = File_Name
File_Name = File_Name & ".txt"

Export_Folder = fn_get_sites_folder(Store_Id,"Export")
File_Full_Name = Export_Folder&File_Name
If FileObject.FileExists(File_Full_Name) Then
   FileObject.DeleteFile(File_Full_Name)
end if
Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
MyFile.Write TxtExport
MyFile.Close
set FileObject = Nothing

fn_redirect "Download_Exported_files.asp?export=newsletter&new="&efile_name

function removeLines(sText)
	sText=nulltoNothing(sText)
    removeLines=replace(replace(replace(replace(sText,vbcrlf,""),vbcr,""),vblf,""),chr(9),"")
end function

function nulltoNothing(sText)
	if not isnull(sText) then
		nulltoNothing = sText
	else
		nulltoNothing = ""
	end if
end function
%>