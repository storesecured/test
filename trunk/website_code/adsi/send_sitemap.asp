<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->

<%

server.scripttimeout=1000

Set xObj = CreateObject("SOFTWING.ASPtear")
on error resume next

set storefields=server.createobject("scripting.dictionary")
sql_select="SELECT store_settings.store_id,site_name,use_domain_name,store_domain FROM store_settings WITH (NOLOCK) WHERE store_id<>101 and trial_version=0 and store_cancel is Null and store_active=1 and overdue_payment=0 and store_id<>101 order by store_id desc"
Call DataGetrows(conn_store,sql_select,storedata,storefields,noRecords)
FOR storerowcounter = 0 TO storefields("rowcount")
    sStore_Id = storedata(storefields("store_id"),storerowcounter)
    Store_name = storedata(storefields("store_name"),storerowcounter)
    Site_Name = storedata(storefields("site_name"),storerowcounter)
    Store_Domain = storedata(storefields("store_domain"),storerowcounter)
    Use_Domain_Name = storedata(storefields("use_domain_name"),storerowcounter)
    
    if Use_Domain_Name and Store_Domain<>"" then
        Site_Name=Store_Domain
    end if
    Site_Name="http://"&Site_Name&"/"
    
    sEncoded_sn = server.URLEncode(Site_Name&"auto_sitemap.xml")
    sEncoded_pd = server.URLEncode(Site_Name&"products.xml")
    
    response.Write "<BR>"&Site_Name
    strResult = xObj.Retrieve("http://www.google.com/webmasters/sitemaps/ping", 1, "sitemap="&sEncoded_sn, "", "")
    strResult = xObj.Retrieve("http://www.google.com/webmasters/sitemaps/ping", 1, "sitemap="&sEncoded_pd, "", "")
    'response.Write "<BR>"&strResult
    strResult = xObj.Retrieve("http://search.yahooapis.com/SiteExplorerService/V1/updateNotification", 1, "appid=StoreSecured&url="&sEncoded_sn, "", "")
    strResult = xObj.Retrieve("http://search.yahooapis.com/SiteExplorerService/V1/updateNotification", 1, "appid=StoreSecured&url="&sEncoded_pd, "", "")
    'response.Write "<BR>"&strResult
next

set storefields=Nothing
set xObj=Nothing



%>