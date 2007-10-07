<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->

<%
on error resume next
server.scripttimeout = 2300

sql_select = "SELECT * from sys_adsi where adsi_server='"&sIP&"' and adsi_completed=0 and adsi_command like 'Delete%'"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
response.write sql_select

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		adsi_options = mydata(myfields("adsi_options"),rowcounter)
		adsi_command = mydata(myfields("adsi_command"),rowcounter)
		adsi_id = mydata(myfields("adsi_id"),rowcounter)
		response.write "<BR>"&adsi_command&"="&adsi_options
		sStatus=0
		select case adsi_command
                       case "Delete_Web"
                            sStatus=Delete_Web (adsi_options)
                       case "Delete_Dir"
                            sStatus=Delete_Dir (adsi_options)
                       case "Delete_Logs"
                            sStatus=Delete_Logs (adsi_options)
                       case "Delete_Site"
                            sStatus=Delete_Site (adsi_options)
                end select
                response.write "sStatus="&sStatus
                if sStatus=1 then
                   sql_update = "update sys_adsi set adsi_completed=1 where adsi_id="&adsi_id
                   response.write "<BR>"&sql_update
                   conn_store.execute sql_update
                end if
	Next
end if

Function Delete_Web (SiteID)
    on error resume next
    Set w3svr = GetObject("IIS://localhost/w3svc")
    w3svr.delete "IIsWebServer", SiteID

    w3svr.SetInfo

    Set w3svr = Nothing
    Delete_Web=1

End Function

function Delete_Site (Del_Store)

	sql_del = "exec wsp_admin_store_delete "&Del_Store
	conn_store.execute sql_del
	Delete_Site=1
	
end function

function Delete_Dir (Del_Store)
	Store_Id=Del_Store

	Set fso = CreateObject("Scripting.FileSystemObject")
	sPath = fn_get_sites_folder(Del_Store,"Base")
    if fso.FolderExists(sPath) then
       fso.DeleteFolder sPath, True
    end if

	Set fso = Nothing
	Delete_Dir=1

end function

function Delete_Logs (Del_Store)

	Set fso = CreateObject("Scripting.FileSystemObject")

        Set fso = CreateObject("Scripting.FileSystemObject")
        if fso.FolderExists("d:\iislogs\www\w3svc"&Del_Store) then
           fso.DeleteFolder "d:\iislogs\www\w3svc"&Del_Store, True
        end if

	Set fso = Nothing
	Delete_Logs=1

end function
%>