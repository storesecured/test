<!--#include virtual="common/connection.asp"-->
<!--#include file="include\sub.asp"-->
<%

server.scripttimeout = 1000

sWWWLog = "d:\iislogs\www\w3svc"
sPerlPath = "c:\Perl\bin\perl.exe"

Set FileObject = CreateObject("Scripting.FileSystemObject")
Logs_Folder = fn_get_code_folder("Logs")
Scheduled_Folder = fn_get_code_folder("Scheduled")
    
Set f = FileObject.GetFolder(Logs_Folder)
Set fc = f.Files
For Each f1 in fc
   if f1.name <> "stats.pl" then
      FileObject.DeleteFile(Logs_Folder&f1.name)
   end if
Next

sNowMinus2= dateadd("d",-2,now())
sNowMinus1= dateadd("d",-1,now())

sYear = right(DatePart("yyyy",sNowMinus1),2)
sMonth = DatePart("m",sNowMinus1)
sDay = DatePart("d",sNowMinus1)

if sMonth < 10 then
     sMonth = "0"&sMonth
end if
if sDay < 10 then
     sDay = "0"&sDay
end if
sDateMinus1 = sYear & sMonth & sDay
    
sYear = right(DatePart("yyyy",sNowMinus2),2)
sMonth = DatePart("m",sNowMinus2)
sDay = DatePart("d",sNowMinus2)

if sMonth < 10 then
     sMonth = "0"&sMonth
end if
if sDay < 10 then
     sDay = "0"&sDay
end if
sDateMinus2 = sYear & sMonth & sDay

sql_select = "SELECT Store_settings.store_id,server,site_name,store_domain,use_domain_name,store_domain2,gmt_offset,skip_hosts,dns_lookup from store_Settings inner join store_statistics on store_settings.store_id=store_statistics.store_id where service_type >= 1 and trial_version=0 and store_cancel is null and overdue_payment<=8 and store_settings.store_id<>101 and (server="&freeserver&" or Server="&paidserver&") Order by Site_Name"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

iTotalRecords=myfields("rowcount")
iQtrRecords=iTotalRecords/4

if noRecords = 0 then
FOR rowcounter= 0 TO myfields("rowcount")
   Store_Id = mydata(myfields("store_id"),rowcounter)
    Store_domain = mydata(myfields("store_domain"),rowcounter)
    Store_domain2 = mydata(myfields("store_domain2"),rowcounter)
    Site_name=mydata(myfields("site_name"),rowcounter)
    Use_Domain_Name = mydata(myfields("use_domain_name"),rowcounter)
    Server_Id= mydata(myfields("server"),rowcounter)
    domainNamewWWW = "www."&Store_domain
    domainNamewoWWW = Replace(Store_domain,"www.","")
    if Store_domain2 <> "" and not isnull(Store_domain2) then
      domainNamewWWW2 = "www."&Store_domain2
      domainNamewoWWW2 = Replace(Store_domain2,"www.","")
    else
      domainNamewWWW2 = ""
      domainNamewoWWW2 = ""
    end if
    Host_Names = Site_Name & " " & Store_domain & " " & domainNamewWWW & " " & domainNamewoWWW& " " & Store_domain2 & " " & domainNamewWWW2 & " " & domainNamewoWWW2
    gmt_offset = mydata(myfields("gmt_offset"),rowcounter)
    skip_hosts = mydata(myfields("skip_hosts"),rowcounter)
    dns_lookup = mydata(myfields("dns_lookup"),rowcounter)
    if dns_lookup then
       dns_lookup=1
    else
       dns_lookup=0
    end if
    on error resume next

    Conf_File = Logs_Folder&"stats."&Site_Name&".conf"

    FileContent = "LogFile="""&sWWWLog&Store_Id&"\ex%YY%MM%DD.log""" & chr(13) & chr(10)
    
    FileContent = FileContent & "LogType=W" & chr(13) & chr(10)
    if gmt_offset <> 0 then
       FileContent = FileContent & "LoadPlugin=""timezone "&gmt_offset&"""" & chr(13) & chr(10)
    end if
    FileContent = FileContent & "LoadPlugin=""tooltips""" & chr(13) & chr(10)
    FileContent = FileContent & "LogFormat=""%time2 %cs-method %cs-uri-stem %cs-username %c-ip %cs-version %cs(User-Agent) %cs(Referer) %virtualname %sc-status %other %sc-bytes %other""" & chr(13) & chr(10)
    FileContent = FileContent & "LogSeparator="" """ & chr(13) & chr(10)
    FileContent = FileContent & "SiteDomain="""&Site_Name&"""" & chr(13) & chr(10)
    FileContent = FileContent & "HostAliases="""&Host_Names&"""" & chr(13) & chr(10)
    FileContent = FileContent & "DNSLookup="&dns_lookup & chr(13) & chr(10)
    FileContent = FileContent & "DirData="""&fn_get_sites_folder(Store_id,"Statistics")&"""" & chr(13) & chr(10)
    FileContent = FileContent & "DirCgi=""/cgi-bin""" & chr(13) & chr(10)
    FileContent = FileContent & "DirIcons=""http://"&Site_Name&"/logs/icon""" & chr(13) & chr(10)
    FileContent = FileContent & "AllowToUpdateStatsFromBrowser=0" & chr(13) & chr(10)
    FileContent = FileContent & "AllowFullYearView=1" & chr(13) & chr(10)
    FileContent = FileContent & "EnableLockForUpdate=0" & chr(13) & chr(10)
    FileContent = FileContent & "DNSStaticCacheFile=""dnscache.txt""" & chr(13) & chr(10)
    FileContent = FileContent & "DNSLastUpdateCacheFile=""dnscachelastupdate.txt"""  & chr(13) & chr(10)
    'FileContent = FileContent & "SkipDNSLookupFor="""" & chr(13) & chr(10)
    FileContent = FileContent & "AllowAccessFromWebToAuthenticatedUsersOnly=0" & chr(13) & chr(10)
    'FileContent = FileContent & "AllowAccessFromWebToFollowingAuthenticatedUsers="""" & chr(13) & chr(10)
    'FileContent = FileContent & "AllowAccessFromWebToFollowingIPAddresses="""" & chr(13) & chr(10)
    FileContent = FileContent & "CreateDirDataIfNotExists=1" & chr(13) & chr(10)
    FileContent = FileContent & "SaveDatabaseFilesWithPermissionsForEveryone=1" & chr(13) & chr(10)
    FileContent = FileContent & "PurgeLogFile=0" & chr(13) & chr(10)
    FileContent = FileContent & "ArchiveLogRecords=0" & chr(13) & chr(10)
    FileContent = FileContent & "KeepBackupOfHistoricFiles=0" & chr(13) & chr(10)
    FileContent = FileContent & "DefaultFile=""store.asp""" & chr(13) & chr(10)
    FileContent = FileContent & "SkipHosts=""68.101.187.204 " & skip_hosts & """" & chr(13) & chr(10)
    'FileContent = FileContent & "SkipUserAgents="""" & chr(13) & chr(10)
    FileContent = FileContent & "SkipFiles=""/httperror.asp /admin_error.asp /error.asp /logs /stats.pl""" & chr(13) & chr(10)
    'FileContent = FileContent & "OnlyHosts="""" & chr(13) & chr(10)
    'FileContent = FileContent & "OnlyUserAgents="""" & chr(13) & chr(10)
    'FileContent = FileContent & "OnlyFiles="""" & chr(13) & chr(10)
    FileContent = FileContent & "NotPageList=""css js class gif jpg jpeg png bmp ico exe dll""" & chr(13) & chr(10)
    FileContent = FileContent & "ValidHTTPCodes=""200 304 302 305 500 206 301 404""" & chr(13) & chr(10)
    FileContent = FileContent & "ValidSMTPCodes=""1 250""" & chr(13) & chr(10)
    FileContent = FileContent & "AuthenticatedUsersNotCaseSensitive=0" & chr(13) & chr(10)
    FileContent = FileContent & "URLNotCaseSensitive=1" & chr(13) & chr(10)
    FileContent = FileContent & "URLWithAnchor=0" & chr(13) & chr(10)
    FileContent = FileContent & "URLQuerySeparators=""?;""" & chr(13) & chr(10)
    FileContent = FileContent & "URLWithQuery=1" & chr(13) & chr(10)
    FileContent = FileContent & "URLWithQueryWithOnlyFollowingParameters=""item_id categ_id parent_ids""" & chr(13) & chr(10)
    FileContent = FileContent & "URLReferrerWithQuery=0" & chr(13) & chr(10)
    FileContent = FileContent & "WarningMessages=1" & chr(13) & chr(10)
    'FileContent = FileContent & "ErrorMessages="""" & chr(13) & chr(10)
    FileContent = FileContent & "DebugMessages=1" & chr(13) & chr(10)
    FileContent = FileContent & "NbOfLinesForCorruptedLog=50" & chr(13) & chr(10)
    'FileContent = FileContent & "WrapperScript="""" & chr(13) & chr(10)
    FileContent = FileContent & "DecodeUA=0" & chr(13) & chr(10)
    FileContent = FileContent & "MiscTrackerUrl=""/js/awstats_misc_tracker.js""" & chr(13) & chr(10)
    FileContent = FileContent & "LevelForRobotsDetection=2" & chr(13) & chr(10)
    FileContent = FileContent & "LevelForBrowsersDetection=2" & chr(13) & chr(10)
    FileContent = FileContent & "LevelForOSDetection=2" & chr(13) & chr(10)
    FileContent = FileContent & "LevelForRefererAnalyze=2" & chr(13) & chr(10)
    FileContent = FileContent & "UseFramesWhenCGI=0" & chr(13) & chr(10)
    FileContent = FileContent & "DetailedReportsOnNewWindows=1" & chr(13) & chr(10)
    FileContent = FileContent & "Expires=0" & chr(13) & chr(10)
    FileContent = FileContent & "MaxRowsInHTMLOutput=1000" & chr(13) & chr(10)
    FileContent = FileContent & "Lang=""auto""" & chr(13) & chr(10)
    FileContent = FileContent & "DirLang=""./lang""" & chr(13) & chr(10)
    FileContent = FileContent & "ShowMenu=U,VPHBLEX" & chr(13) & chr(10)
    FileContent = FileContent & "ShowMonthStats=UVPHB" & chr(13) & chr(10)
    FileContent = FileContent & "ShowDaysOfMonthStats=VPHB" & chr(13) & chr(10)
    FileContent = FileContent & "ShowDaysOfWeekStats=PHB" & chr(13) & chr(10)
    FileContent = FileContent & "ShowHoursStats=PHB" & chr(13) & chr(10)
    FileContent = FileContent & "ShowDomainsStats=PHB" & chr(13) & chr(10)
    FileContent = FileContent & "ShowHostsStats=PHBL" & chr(13) & chr(10)
    FileContent = FileContent & "ShowAuthenticatedUsers=0" & chr(13) & chr(10)
    FileContent = FileContent & "ShowRobotsStats=HBL" & chr(13) & chr(10)
    FileContent = FileContent & "ShowSessionsStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "ShowPagesStats=PBEX" & chr(13) & chr(10)
    FileContent = FileContent & "ShowFileTypesStats=HB" & chr(13) & chr(10)
    FileContent = FileContent & "ShowFileSizesStats=0" & chr(13) & chr(10)
    FileContent = FileContent & "ShowOSStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "ShowBrowsersStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "ShowScreenSizeStats=0" & chr(13) & chr(10)
    FileContent = FileContent & "ShowOriginStats=PH" & chr(13) & chr(10)
    FileContent = FileContent & "ShowKeyphrasesStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "ShowKeywordsStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "ShowMiscStats=a" & chr(13) & chr(10)
    FileContent = FileContent & "ShowHTTPErrorsStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "ShowSMTPErrorsStats=0" & chr(13) & chr(10)
    FileContent = FileContent & "ShowClusterStats=0" & chr(13) & chr(10)
    FileContent = FileContent & "AddDataArrayMonthStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "AddDataArrayShowDaysOfMonthStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "AddDataArrayShowDaysOfWeekStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "AddDataArrayShowHoursStats=1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfDomain = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitDomain  = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfHostsShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitHost    = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfLoginShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitLogin   = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfRobotShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitRobot   = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfPageShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitFile    = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfOsShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitOs      = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfBrowsersShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitBrowser = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfScreenSizesShown = 5" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitScreenSize = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfRefererShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitRefer   = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfKeyphrasesShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitKeyphrase = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfKeywordsShown = 10" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitKeyword = 1" & chr(13) & chr(10)
    FileContent = FileContent & "MaxNbOfEMailsShown = 20" & chr(13) & chr(10)
    FileContent = FileContent & "MinHitEMail   = 1" & chr(13) & chr(10)
    FileContent = FileContent & "FirstDayOfWeek=0" & chr(13) & chr(10)
    'FileContent = FileContent & "ShowFlagLinks=""""" & chr(13) & chr(10)
    FileContent = FileContent & "ShowLinksOnUrl=1" & chr(13) & chr(10)
    'FileContent = FileContent & "UseHTTPSLinkForUrl=""""" & chr(13) & chr(10)
    FileContent = FileContent & "MaxLengthOfURL=70" & chr(13) & chr(10)
    FileContent = FileContent & "LinksToWhoIs=""http://www.whois.net/search.cgi2?str=""" & chr(13) & chr(10)
    FileContent = FileContent & "LinksToIPWhoIs=""http://ws.arin.net/cgi-bin/whois.pl?queryinput=""" & chr(13) & chr(10)
    'FileContent = FileContent & "HTMLHeadSection=""""" & chr(13) & chr(10)
    'FileContent = FileContent & "HTMLEndSection=""""" & chr(13) & chr(10)
    FileContent = FileContent & "Logo=""spacer.gif""" & chr(13) & chr(10)
    FileContent = FileContent & "LogoLink=""http://awstats.sourceforge.net""" & chr(13) & chr(10)
    FileContent = FileContent & "BarWidth   = 260" & chr(13) & chr(10)
    FileContent = FileContent & "BarHeight  = 90" & chr(13) & chr(10)
    'FileContent = FileContent & "StyleSheet=""""" & chr(13) & chr(10)
    FileContent = FileContent & "color_Background=""FFFFFF""" & chr(13) & chr(10)
    FileContent = FileContent & "color_TableBGTitle=""CCCCDD""" & chr(13) & chr(10)
    FileContent = FileContent & "color_TableTitle=""000000""" & chr(13) & chr(10)
    FileContent = FileContent & "color_TableBG=""CCCCDD""" & chr(13) & chr(10)
    FileContent = FileContent & "color_TableRowTitle=""FFFFFF""" & chr(13) & chr(10)
    FileContent = FileContent & "color_TableBGRowTitle=""ECECEC""" & chr(13) & chr(10)
    FileContent = FileContent & "color_TableBorder=""ECECEC""" & chr(13) & chr(10)
    FileContent = FileContent & "color_text=""000000""" & chr(13) & chr(10)
    FileContent = FileContent & "color_textpercent=""606060""" & chr(13) & chr(10)
    FileContent = FileContent & "color_titletext=""000000""" & chr(13) & chr(10)
    FileContent = FileContent & "color_weekend=""EAEAEA""" & chr(13) & chr(10)
    FileContent = FileContent & "color_link=""0011BB""" & chr(13) & chr(10)
    FileContent = FileContent & "color_hover=""605040""" & chr(13) & chr(10)
    FileContent = FileContent & "color_u=""FFB055""" & chr(13) & chr(10)
    FileContent = FileContent & "color_v=""F8E880""" & chr(13) & chr(10)
    FileContent = FileContent & "color_p=""4477DD""" & chr(13) & chr(10)
    FileContent = FileContent & "color_h=""66F0FF""" & chr(13) & chr(10)
    FileContent = FileContent & "color_k=""2EA495""" & chr(13) & chr(10)
    FileContent = FileContent & "color_s=""8888DD""" & chr(13) & chr(10)
    FileContent = FileContent & "color_e=""CEC2E8""" & chr(13) & chr(10)
    FileContent = FileContent & "color_x=""C1B2E2""" & chr(13) & chr(10)
    
    Set MyFile = FileObject.OpenTextFile(Conf_File, 2,true)
    MyFile.Write FileContent
    MyFile.Close

    on error goto 0

	FileContent = vbcrlf &vbcrlf & "mkdir "&sWWWLog&Store_id&"\"
	FileContent = FileContent & vbcrlf & """d:\program files\awstats\tools\logresolvemerge.pl"" \\10.235.158.133\d$\iislogs\www\w3svc"&Store_Id&"\ex"&sDateMinus2&".log \\10.235.158.134\d$\iislogs\www\w3svc"&Store_Id&"\ex"&sDateMinus2&".log > d:\iislogs\www\w3svc"&Store_Id&"\ex"&sDateMinus2&".log"
	FileContent = FileContent & vbcrlf & sPerlPath&" """&Logs_Folder&"stats.pl"" -config="&Site_Name&" -update -logfile="&sWWWLog&Store_id&"\ex"&sDateMinus2&".log"
	FileContent = FileContent & vbcrlf & """d:\program files\awstats\tools\logresolvemerge.pl"" \\10.235.158.133\d$\iislogs\www\w3svc"&Store_Id&"\ex"&sDateMinus1&".log \\10.235.158.134\d$\iislogs\www\w3svc"&Store_Id&"\ex"&sDateMinus1&".log > d:\iislogs\www\w3svc"&Store_Id&"\ex"&sDateMinus1&".log"
	FileContent = FileContent & vbcrlf & sPerlPath&" """&Logs_Folder&"stats.pl"" -config="&Site_Name&" -update -logfile="&sWWWLog&Store_id&"\ex"&sDateMinus1&".log"

    if rowcounter>(iQtrRecords*3) then
     	sFileName="stats4.bat"
    elseif rowcounter>(iQtrRecords*2) then
    		sFileName="stats3.bat"
    elseif rowcounter>(iQtrRecords) then
		sFileName="stats2.bat"
	else
		sFileName="stats1.bat"
	end if
    on error resume next

    Set MyFile = FileObject.OpenTextFile(Scheduled_Folder&sFileName, 8,true)
    MyFile.Write FileContent
    MyFile.Close
Next
End if

set myfields = Nothing
set FileObject = Nothing

%>

