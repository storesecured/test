<!--#include file="Global_Settings.asp"-->
<!--#include file="bigdownload.asp"-->

<% 

server.ScriptTimeout = 40000

File_Location =checkStringForQ(Replace(Replace((Request.QueryString("File_Location")),"../",""),"..\",""))

Root_Folder = fn_get_sites_folder(Store_Id,"Root")
File_Full_Name = Root_Folder&"\"&File_Location
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If (objFSO.FileExists(File_Full_Name))  Then
	SendFileByBlocks File_Full_Name, 65535
else
	response.write "File does not exist, please email store administrator for help."
end if
Set objFSO = Nothing

%>
