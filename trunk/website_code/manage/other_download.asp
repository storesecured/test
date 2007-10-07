<!--#include file="Global_Settings.asp"-->
<!--#include file="bigdownload.asp"-->

<% 

server.ScriptTimeout = 40000

File_Location =checkStringForQ(Replace(Replace((Request.QueryString("File_Location")),"../",""),"..\",""))

Manage_Path = fn_get_code_folder("Manage")
Destination_File_Name = Manage_Path&File_Location
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If (objFSO.FileExists(Destination_File_Name))  Then
	SendFileByBlocks Destination_File_Name, 65535
else
	response.write "File does not exist, please email store administrator for help."
end if
Set objFSO = Nothing

%>
