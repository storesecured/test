<!--#include file="header_noview.asp"-->
<%
on error goto 0

server.ScriptTimeout = 40000
File_Location =checkStringForQ(request.querystring("File_Location"))
sql_select = "exec wsp_purchase_download "&store_id&","&cid&",'"&File_Location&"';"
fn_print_debug sql_select
set myfieldsloc=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydataloc,myfieldsloc,noRecordsloc)
if noRecordsloc <> 0 then
    fn_error "You do not have permission to download this file."
end if
set myfieldsloc=Nothing
 
Download_Folder = fn_get_sites_folder(Store_Id,"Download")   
Download_File_Name = Download_Folder&File_Location
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If (objFSO.FileExists(Download_File_Name))  Then
	Set objFSO = Nothing
    fn_redirect Switch_Name&"include/bigdownload.asp?File_Location="&File_Location
else
	Set objFSO = Nothing
	fn_error "File does not exist, please email store administrator at "&store_email&" for assistance."
end if

%>

