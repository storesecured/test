<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<%
on error resume next

Set fso = CreateObject("Scripting.FileSystemObject")
Store_Id=8588
Old_Group = Round(Store_Id/1000)
sDrive="c"

Root_Folder=fn_get_sites_folder(Store_Id,"Root")
Base_Folder=fn_get_sites_folder(Store_Id,"Base")
Statistics_Folder=fn_get_sites_folder(Store_Id,"Statistics")
response.Write Base_Folder
if fso.FolderExists(Base_Folder)=false then
    set f=fso.CreateFolder(Base_Folder)
    set f=Nothing
end if
if fso.FolderExists(Root_Folder)=false then
    set f=fso.CreateFolder(Root_Folder)
    set f=Nothing
end if
if fso.FolderExists(Statistics_Folder)=false then
    set f=fso.CreateFolder(Statistics_Folder)
    set f=Nothing
end if

fso.CopyFolder sDrive&":\sites\group_"&Old_Group&"\files_"&store_Id&"\*",Root_Folder
fso.CopyFile sDrive&":\sites\group_"&Old_Group&"\files_"&store_Id&"\*",Root_Folder
fso.CopyFolder sDrive&":\files\group_"&Old_Group&"\files_"&store_Id&"\*",Base_Folder
'fso.CopyFolder sDrive&":\iislogs\www\w3svc\"&Store_Id,fn_get_sites_folder(Store_Id,"Log")
fso.CopyFile Root_Folder&"log_data\*.*",Statistics_Folder
fso.DeleteFolder Root_Folder&"aspnet_client"
fso.DeleteFolder Base_Folder&"froogle"
fso.DeleteFolder Base_Folder&"ftp"
fso.DeleteFolder Root_Folder&"log_data"

set fso=Nothing

Set objNew = GetObject("IIS://localhost/W3SVC/"&Store_id)
objNew.SetInfo

strObjName = "IIS://localhost/W3SVC/" & Store_id & "/Root"
Set IIsWebVDirRootObj = GetObject(strObjName)
IIsWebVDirRootObj.Put "Path", Root_Folder

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "common"
IIsWebVDirRootObj.SetInfo

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "include"
IIsWebVDirRootObj.SetInfo

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "logs\log_data"
IIsWebVDirRootObj.SetInfo

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "logs"
IIsWebVDirRootObj.SetInfo

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "store_engine"
IIsWebVDirRootObj.SetInfo

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "images\images_themes"
IIsWebVDirRootObj.SetInfo

IIsWebVDirRootObj.Delete "IIsWebVirtualDir", "images\images_"&store_id
IIsWebVDirRootObj.SetInfo

set IIsWebVDirRootObj= nothing

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root")
objRoot.Path = fn_get_sites_folder(SiteID,"Root")
objRoot.AccessRead=True
objRoot.AccessScript=True
objRoot.SetInfo

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root/common")
objRoot.Path = fn_get_code_folder("Common")
objRoot.SetInfo

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root/include")
objRoot.Path = fn_get_code_folder("Include")
objRoot.SetInfo

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root/store_engine")
objRoot.Path = fn_get_code_folder("Store_Engine")
objRoot.SetInfo

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root/images/images_themes")
objRoot.Path = fn_get_code_folder("Images_Themes")
objRoot.SetInfo

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root/images/images_"&store_id)
objRoot.Path = fn_get_sites_folder(store_id,"Images")
objRoot.SetInfo

Set objRoot = objNew.Create ("IIsWebVirtualDir", "Root/logs")
objRoot.Path = fn_get_code_folder("Logs")
objRoot.AccessExecute=True
objRoot.DefaultDoc="stats.pl"
objRoot.SetInfo

set objNew = nothing
set objRoot = nothing
 %>