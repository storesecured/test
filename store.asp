<!--#include virtual="common/connection.asp"-->

<%
sLocalIP = Request.ServerVariables ("LOCAL_ADDR")
if sLocalIP=sPortalIP then
    'don't setup the sites if this is the portal server
    response.Write "no sites setup, this is the portal server"
    response.end
end if
iLocalIP = fn_convert_ip_to_bigint(sLocalIP)

sName = Request.ServerVariables("HTTP_HOST")

'find all of the stores which have been newly activated as paid stores and are not yet moved over
sql_select = "SELECT store_id from store_domains where Store_Id<>101 and store_url ='"&sName&"' and redirect_only=0" 

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
'if there are no records it means a store should not exist, otherwise if there are then create it
if noRecords = 0 then
    FOR rowcounter= 0 TO myfields("rowcount")
	    if response.isclientconnected=0 then
		    response.end
        end if
	Store_Id = mydata(myfields("store_id"),rowcounter)

    
	Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP.6.0")
        xmlHttp.setOption 2, 13056
        xmlhttp.Open "POST","http://localhost/create_site.asp?Store_id="&Store_id,false
        xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
        xmlHttp.Send(Post_String)
        
        strResult = xmlHttp.responseText     
        xmlHttp.abort()
        set xmlHttp = Nothing
        
        if err.number<>0 then
	        response.Write "There was an error."
        else
            'redirect to site now that it is created
            response.redirect "http://"&sName
        end if
    Next
else
    response.Write "Could not find a matching store"
end if

%>


</body>
</html>


