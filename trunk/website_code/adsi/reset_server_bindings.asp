<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<html>

<head>
<!-- #INCLUDE FILE="functions2.asp" -->



<%

on error goto 0

server.scripttimeout=4800

'find all of the stores which have been newly activated as paid stores and are not yet moved over
sql_select = "SELECT Store_Id,site_name,secure_name,store_domain,store_domain2 from Store_Settings where store_id=8588"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
	FOR rowcounter= 0 TO myfields("rowcount")
		Store_Id = mydata(myfields("store_id"),rowcounter)
		Store_Domain = mydata(myfields("store_domain"),rowcounter)
		Store_Domain2 = mydata(myfields("store_domain2"),rowcounter)
		Site_Name = mydata(myfields("site_name"),rowcounter)
        Secure_Name = mydata(myfields("secure_name"),rowcounter)

        sDomainList=Site_Name&","&Secure_Name&","&Store_Domain&","&Store_Domain2
        sCompleted = fn_server_bindings (Store_Id, sDomainList)
	Next
end if


response.write "Done"
%>
