<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%
server.scripttimeout=120

sql_select = "select distinct store_domain,store_domain2,overdue_payment from store_settings"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

Set DNS = Server.CreateObject("ASPDNS.DNSLookup")

Set fso = CreateObject("Scripting.FileSystemObject")
folderspec = Mail_Path

Set fsub = fso.GetFolder(folderspec)
Set fc = fsub.subfolders
For Each f2 in fc
    f2=lcase(f2)
    sFoldername=replace(f2,folderspec,"")
    iFound=0
    FOR rowcounter= 0 TO myfields("rowcount")
        overdue_payment = lcase(mydata(myfields("overdue_payment"),rowcounter))
        store_domain = lcase(mydata(myfields("store_domain"),rowcounter))
        store_domain2 = lcase(mydata(myfields("store_domain2"),rowcounter))  
        if store_domain<>"" then
            store_domain = replace(store_domain,"www.","")
        end if
        if store_domain2<>"" then
            store_domain2 = replace(store_domain2,"www.","")  
        end if
        if sFoldername = store_domain or sFoldername = store_domain2 then
            iFound=1
            if overdue_payment>30 then
            	response.write "<BR>**OVERDUE "&overdue_payment&" ** "&sFoldername
            end if
            exit for
        end if
    next
    if iFound=0 then
        'Site_Ip = DNS.GetIPFromName("http://"&f2)
        if Site_IP<>"69.20.120.19" and Site_Ip<>"207.97.238.253" then
            response.write "<BR>**REMOVE ** "&sFoldername
        else
            response.write "<BR>**MAYBE REMOVE ** "&sFoldername
        end if
    else
        'response.write "<BR>**GOOD ** "&sFoldername
    end if

Next

Set DNS=Nothing
Set fso=nothing
set myfields=nothing

%>