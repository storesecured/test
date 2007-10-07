<!--#include virtual="common/connection.asp"-->
<%
on error resume next

server.scripttimeout = 100
folderspec = "d:\iislogs\www\"

Dim fso, f, f1, fc
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set ff = f.SubFolders
fileCnt = f.Files.Count


tfiles = 0
sDate=now()

For Each f1 in ff
    sName= f1.name
    Set fsub = fso.GetFolder(folderspec&sName)
    Set fc = fsub.Files
    For Each f2 in fc
      if datediff("m",f2.datecreated,sDate)>=2 then
              fso.DeleteFile folderspec&sName&"\"&f2.name
              tfiles=tfiles+1
      end if
    next
next
response.write "deleted "&tfiles&" files"

set fc=nothing
set ff=nothing
set f=nothing
set fso=nothing

%>