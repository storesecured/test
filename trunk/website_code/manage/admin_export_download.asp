<!--#include file="Global_Settings.asp"-->
<!--#include file="bigdownload.asp"-->

<% 

server.ScriptTimeout = 40000

File_Location =checkStringForQ(Replace(Replace((Request.QueryString("File_Location")),"../",""),"..\",""))


Destination_Folder = fn_get_sites_folder(Store_Id,"Export")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If (objFSO.FileExists(Destination_Folder))  Then
	SendFileByBlocks Destination_Folder, 65535
else
	fn_error "File does not exist, please email store administrator for help."
end if
Set objFSO = Nothing

%>
