<!--#include file="Global_Settings.asp"-->
<!--#include file="bigdownload.asp"-->

<% 

server.ScriptTimeout = 40000

File_Location =checkStringForQ(Replace(Replace((Request.QueryString("File_Location")),"../",""),"..\",""))


DestinationPath = File_Location
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If (objFSO.FileExists(DestinationPath))  Then
	SendFileByBlocks DestinationPath, 65535
	'mySmartUpload.DownloadFile(DestinationPath)
else
	response.write "File does not exist, please email store administrator for help."
end if
Set objFSO = Nothing
%>
