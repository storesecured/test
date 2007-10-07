<!--#include file="include/header.asp"-->

<%

Set FileObject = CreateObject("Scripting.FileSystemObject")
Root_Folder = fn_get_sites_folder(Store_Id,"Root")
File_Full_Name=Root_Folder&"redirect.txt"

If FileObject.FileExists(File_Full_Name) Then
    fn_print_debug "in file exists"
    'check for redirect path
    Set MyFile = FileObject.OpenTextFile(File_Full_Name, 1)
    Do While MyFile.AtEndOfStream <> True
        ReadLineTextFile = MyFile.ReadLine
        sSplitName=split(ReadLineTextFile,"|")
        sOldName=sSplitName(0)
        if sOldName=ReturnTo_Orig then
            'found a match so redirect
            sNewName=sSplitName(1)
            MyFile.Close
            set FileObject=Nothing
            fn_redirect_perm sNewName
        end if
    Loop
    MyFile.Close
    if ReadAll<>"" then
    
    end if
end if
set FileObject=Nothing

  
response.status="404 not found" %>
<TABLE width="100%">


    <TR><TD>The requested page could not be found.<BR><BR>
    <a href=<%= Site_Name %> class=link>Please click here to go to the homepage.</a></TD></TR>


</TABLE>

<!--#include file="include/footer.asp"-->
