<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->

<%
server.scripttimeout = 8000

if request.servervariables("Remote_Addr") = request.servervariables("Local_Addr") then
    FTP_TRANSFER_TYPE_ASCII = 1
    FTP_TRANSFER_TYPE_BINARY = 2
    Set FtpConn = Server.CreateObject("AspInet.FTP")
    Set fso = CreateObject("Scripting.FileSystemObject")

    sql_select_items = "select store_id,froggle_filename,froggle_username,froggle_password,froggle_server,froggle_submit from store_external where froggle_enable=1 and (froggle_submit Is Null OR feeds_last_created>froggle_submit)"
    set myfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select_items,mydata,myfields,noRecords)
    response.write sql_select_items

    i = 0
    if noRecords = 0 then
        FOR rowcounter= 0 TO myfields("rowcount")
            iStore_ID = mydata(myfields("store_id"),rowcounter)
            	
	        Filename = mydata(myfields("froggle_filename"),rowcounter)
	        Username = mydata(myfields("froggle_username"),rowcounter)
	        Password = mydata(myfields("froggle_password"),rowcounter)
	        sServer = mydata(myfields("froggle_server"),rowcounter)
        	
            Froogle_Filename = fn_get_sites_folder(iStore_id,"Root")&"froogle.txt"
            if fso.fileexists(Froogle_Filename) then
                
                sError = FtpConn.FTPPutFile(sServer, Username, Password, Filename&".txt", Froogle_Filename, FTP_TRANSFER_TYPE_ASCII)
                if sError then
                    Response.Write "<p>FTP upload Success...<br>"
                    sql_update = "Update Store_External set Froggle_Submit='"&Now()&"' WHERE Store_ID = " & iStore_ID
                    conn_store.Execute sql_update
                else
                    Response.Write "<BR><BR><br>"  & sError & "-"&FtpConn.lasterror
                    Response.Write "<p>FTP upload Failed...<br>" & Store_Id&"<BR>"& Froogle_Filename
                    response.write "<BR>directory="&sServer&sDirectory
                    response.write "<BR>filename="&Filename
                    response.write "<BR>Username="&Username
                    response.write "<BR>Password="&Password
                    response.write "<BR>sServerPath="&Froogle_Filename&"<BR>"
	            end if
            end if

            if not(response.isclientconnected) then
                response.end
            end if

        next
    end if
end if

%>



