<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%

on error goto 0
server.scripttimeout = 4800

if request.servervariables("Remote_Addr") = request.servervariables("Local_Addr") then

Ftp_Folder = fn_get_code_folder("Ftp")
Set FileObject = CreateObject("Scripting.FileSystemObject")
Set f = FileObject.GetFolder(Ftp_Folder)
Set fc = f.Files
For Each f1 in fc
   FileObject.DeleteFile(Ftp_Folder&f1.name)
Next

on error resume next
FileObject.DeleteFile(FTP_User_File)
on error goto 0

Ftp_Folder = fn_get_code_folder("Ftp")

CustomerContent = ("AllowChangePassword=0" & vbcrlf)
CustomerContent = CustomerContent & ("EnableGroup=1" & vbcrlf)
CustomerContent = CustomerContent & ("GroupName=customers" & vbcrlf)
CustomerContent = CustomerContent & ("AllowNoop=1" & vbcrlf)
CustomerContent = CustomerContent & ("Home-Ip=-= All IP Homes =-" & vbcrlf)
CustomerContent = CustomerContent & ("RelativePath=1" & vbcrlf)
CustomerContent = CustomerContent & ("TimeOut=600" & vbcrlf)
CustomerContent = CustomerContent & ("EnableMaxConPerIP=0" & vbcrlf)
CustomerContent = CustomerContent & ("MaxConPerIp=0" & vbcrlf)
CustomerContent = CustomerContent & ("MaxUsers=0" & vbcrlf)
CustomerContent = CustomerContent & ("RatioMethod=0" & vbcrlf)
CustomerContent = CustomerContent & ("RatioUp=1" & vbcrlf)
CustomerContent = CustomerContent & ("RatioDown=1" & vbcrlf)
CustomerContent = CustomerContent & ("RatioCredit=0" & vbcrlf)
CustomerContent = CustomerContent & ("QuotaEnabled=1" & vbcrlf)
CustomerContent = CustomerContent & ("MaxSpeedEnabled=0" & vbcrlf)
CustomerContent = CustomerContent & ("MaxSpeedRcv=4192" & vbcrlf)
CustomerContent = CustomerContent & ("MaxSpeedSnd=4192" & vbcrlf)
CustomerContent = CustomerContent & ("AddLinks=1" & vbcrlf)
CustomerContent = CustomerContent & ("TreatLinksAs=0" & vbcrlf)
CustomerContent = CustomerContent & ("AddLinkFromFile=1" & vbcrlf)
CustomerContent = CustomerContent & ("AddHomeLink=0" & vbcrlf)
CustomerContent = CustomerContent & ("Hide Hidden Files=1" & vbcrlf)
CustomerContent = CustomerContent & ("DefaultGroupRatioCredit=0" & vbcrlf)
CustomerContent = CustomerContent & ("DefaultGroupQuotaCredit=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_Login=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_LastLogin=12/30/1899" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_LastIP=Unknown" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_KBUp=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_KBDown=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_FilesUp=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_FilesDown=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_FailedUp=0" & vbcrlf)
CustomerContent = CustomerContent & ("Stat_FailedDown=0" & vbcrlf)

AdminContent = ("AllowChangePassword=0" & vbcrlf)
AdminContent = AdminContent & ("Home-Ip=-= All IP Homes =-" & vbcrlf)
AdminContent = AdminContent & ("RelativePath=1" & vbcrlf)
AdminContent = AdminContent & ("TimeOut=600" & vbcrlf)
AdminContent = AdminContent & ("MaxConPerIp=2" & vbcrlf)
AdminContent = AdminContent & ("MaxUsers=0" & vbcrlf)
AdminContent = AdminContent & ("RatioMethod=0" & vbcrlf)
AdminContent = AdminContent & ("RatioUp=1" & vbcrlf)
AdminContent = AdminContent & ("RatioDown=1" & vbcrlf)
AdminContent = AdminContent & ("RatioCredit=0" & vbcrlf)
AdminContent = AdminContent & ("MaxSpeedRcv=512" & vbcrlf)
AdminContent = AdminContent & ("MaxSpeedSnd=512" & vbcrlf)
AdminContent = AdminContent & ("QuotaCurrent=0" & vbcrlf)
AdminContent = AdminContent & ("QuotaMax=0" & vbcrlf)
AdminContent = AdminContent & ("AddLinks=1" & vbcrlf)
AdminContent = AdminContent & ("TreatLinksAs=0" & vbcrlf)
AdminContent = AdminContent & ("ResolveLNK=1" & vbcrlf)
AdminContent = AdminContent & ("Stat_Login=63" & vbcrlf)
AdminContent = AdminContent & ("Stat_LastLogin=12/30/1899" & vbcrlf)
AdminContent = AdminContent & ("Stat_LastIP=Unknown" & vbcrlf)
AdminContent = AdminContent & ("Stat_KBUp=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_KBDown=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FilesUp=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FilesDown=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FailedUp=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FailedDown=0" & vbcrlf)
AdminContent = AdminContent & ("LinksFile="&Ftp_Folder&"home.txt" & vbcrlf)
AdminContent = AdminContent & ("AddLinkFromFile=1" & vbcrlf)
AdminContent = AdminContent & ("Stat_Login=63" & vbcrlf)
AdminContent = AdminContent & ("Stat_LastLogin=12/30/1899" & vbcrlf)
AdminContent = AdminContent & ("Stat_LastIP=Unknown" & vbcrlf)
AdminContent = AdminContent & ("Stat_KBUp=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_KBDown=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FilesUp=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FilesDown=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FailedUp=0" & vbcrlf)
AdminContent = AdminContent & ("Stat_FailedDown=0" & vbcrlf)

EmployeeContent = ("Dir0=D:\home_melanie" & vbcrlf)
EmployeeContent = EmployeeContent & ("Attr0=R----LS-" & vbcrlf)
EmployeeContent = EmployeeContent & ("Dir1=d:\" & vbcrlf)
EmployeeContent = EmployeeContent & ("Attr1=R----LS-" & vbcrlf)

FileContent = "[blac6789]" & vbcrlf
FileContent = FileContent & ("Login=blac6789" & vbcrlf)
FileContent = FileContent & ("Pass=tom237" & vbcrlf)
FileContent = FileContent & AdminContent
FileContent = FileContent & ("Dir0=D:\home_melanie" & vbcrlf)
FileContent = FileContent & ("Attr0=R----LS-" & vbcrlf)
FileContent = FileContent & ("Dir1=C:\" & vbcrlf)
FileContent = FileContent & ("Attr1=RWDAMLSK" & vbcrlf)
FileContent = FileContent & ("Dir2=d:\" & vbcrlf)
FileContent = FileContent & ("Attr2=RWDAMLSK" & vbcrlf)

FileContent = FileContent & AdminContent
FileContent = FileContent & (vbcrlf)

FileContent = FileContent& "[jason]" & vbcrlf
FileContent = FileContent & ("Login=jason" & vbcrlf)
FileContent = FileContent & ("Pass=nadin2005" & vbcrlf)
FileContent = FileContent & AdminContent
FileContent = FileContent & ("Dir0=D:\home_melanie" & vbcrlf)
FileContent = FileContent & ("Attr0=R----LS-" & vbcrlf)
FileContent = FileContent & ("Dir1=C:\" & vbcrlf)
FileContent = FileContent & ("Attr1=RWDAMLSK" & vbcrlf)
FileContent = FileContent & ("Dir2=d:\" & vbcrlf)
FileContent = FileContent & ("Attr2=RWDAMLSK" & vbcrlf)

FileContent = FileContent & AdminContent
FileContent = FileContent & (vbcrlf)


FileContent = FileContent & ("[abrar]" & vbcrlf)
FileContent = FileContent & ("Login=abrar" & vbcrlf)
FileContent = FileContent & ("Pass=sohail" & vbcrlf)
FileContent = FileContent & AdminContent & EmployeeContent

FileContent = FileContent & ("Dir0="&fn_get_code_folder("Sales") & vbcrlf)
FileContent = FileContent & ("Attr0=RWDAMLSK" & vbcrlf)
User_File = Ftp_Folder&"abrar.txt"
Set MyFile = FileObject.OpenTextFile(User_File, 8,true)
MyFile.Write FileContent1
MyFile.Close


sql_select = "SELECT store_id,store_user_id,store_password,service_type,server,additional_storage from store_Settings where service_type >= 1 and store_id<>101 and store_cancel is null and overdue_payment<=8 Order by Store_Id"
response.write sql_select
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

if noRecords = 0 then
FOR rowcounter= 0 TO myfields("rowcount")
    Store_Id = mydata(myfields("store_id"),rowcounter)
    store_user_id = mydata(myfields("store_user_id"),rowcounter)
    store_password = mydata(myfields("store_password"),rowcounter)
    service_type = mydata(myfields("service_type"),rowcounter)
    additional_storage = mydata(myfields("additional_storage"),rowcounter)
    server_id = mydata(myfields("server"),rowcounter)

    if server_id=1 or server_id=2 then
       sInstalledDrive="y"
    elseif server_id=3 or server_id=4 then
       sInstalledDrive="s"
    elseif server_id=5 or server_id=6 then
       sInstalledDrive="k"
    elseif server_id=7 or server_id=8 then
       sInstalledDrive="g"
    elseif server_id=9 or server_id=10 then
       sInstalledDrive="d"
    end if

    on error goto 0
    sType="paid"

    if Service_Type = 0 then
		available = 5
		sType="free"
	elseif Service_Type = 3 then
		available = 50
	elseif Service_Type = 5 then
		available = 100
	elseif Service_Type = 7 then
		available = 250
	elseif Service_Type = 9 or Service_Type = 10 then
		available = 500
	else
		available = 1000
	end if

	if Additional_Storage > 0 then
		available = available + (100 * Additional_Storage)
	end if
	available = available * 1024 * 1024

         total_size=0
    Base_Folder = fn_get_sites_folder(Store_Id,"Base")
    Root_Folder = fn_get_sites_folder(Store_Id,"Root")
    Ftp_Links_Folder = fn_get_code_folder("Ftp")
    Log_Folder = fn_get_sites_folder(Store_Id,"Log")
    Export_Folder = fn_get_sites_folder(Store_Id,"Export")
    Download_Folder = fn_get_sites_folder(Store_Id,"Download")
    Key_Folder = fn_get_sites_folder(Store_Id,"Key")
    Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")
    
	FileContent = FileContent & ("["&store_id&"]" & vbcrlf)
	FileContent = FileContent & ("Login="&store_user_id & vbcrlf)
	FileContent = FileContent & ("Pass="&store_password & vbcrlf)
	FileContent = FileContent & ("QuotaCurrent="&total_size & vbcrlf)
	FileContent = FileContent & ("QuotaMax="&available & vbcrlf)
	FileContent = FileContent & ("LinksFile="&Ftp_Links_Folder&Store_Id&".txt" & vbcrlf)
	FileContent = FileContent & CustomerContent
    FileContent = FileContent & ("Dir0="&Base_Folder & vbcrlf)
    FileContent = FileContent & ("Attr0=-----L--" & vbcrlf)
    FileContent = FileContent & ("Dir1="&Root_Folder & vbcrlf)
    FileContent = FileContent & ("Attr1=RWDAMLSK" & vbcrlf)
    FileContent = FileContent & ("Dir2="&Export_Folder & vbcrlf)
    FileContent = FileContent & ("Attr2=RWDAMLSK" & vbcrlf)
    FileContent = FileContent & ("Dir3="&Log_Folder & vbcrlf)
    FileContent = FileContent & ("Attr3=R----L--" & vbcrlf)
    FileContent = FileContent & ("Dir4="&Download_Folder & vbcrlf)
    FileContent = FileContent & ("Attr4=RWDAMLSK" & vbcrlf)
    FileContent = FileContent & ("Dir5="&Key_Folder & vbcrlf)
    FileContent = FileContent & ("Attr5=RWDAMLSK" & vbcrlf)
    FileContent = FileContent & ("Dir6="&Upload_Folder & vbcrlf)
    FileContent = FileContent & ("Attr6=RWDAMLSK" & vbcrlf)
    
    FileContent1 = "To Home Directory | ~" & vbcrlf
    FileContent1 = FileContent1 & ("Hidden Admin Files | "& Base_Folder & vbcrlf)
    FileContent1 = FileContent1 & ("Raw Access Logs | "& Log_Folder & vbcrlf)
    FileContent1 = FileContent1 & ("Web Site Files | "& Root_Folder & vbcrlf)
   
    FileContent = FileContent & (vbcrlf)

	Set MyFile = FileObject.OpenTextFile(fn_get_code_folder("Ftp")&Store_Id&".txt", 8,true)
	MyFile.Write FileContent1
	MyFile.Close
Next
    Set MyFile = FileObject.OpenTextFile(FTP_User_File, 8,true)
    MyFile.Write FileContent
    MyFile.Close
End if

set myfields = Nothing
set FileObject = Nothing

end if

%>


