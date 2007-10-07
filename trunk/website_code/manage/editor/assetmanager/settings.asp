<%
Dim arrBaseFolder
Redim arrBaseFolder(4)
Dim arrBaseName
Redim arrBaseName(4)

Dim bReturnAbsolute
bReturnAbsolute=false

arrBaseFolder(0)=fn_get_sites_folder(Store_id,"Root")
arrBaseName(0)="Store Files"

arrBaseFolder(1)=""
arrBaseName(1)=""

arrBaseFolder(2)=""
arrBaseName(2)=""

arrBaseFolder(3)=""
arrBaseName(3)=""
%>