<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
		 
<%

			sOpenPath = replace(replace(request.querystring("Open"),"#","/"),"../","")
			sNewPath = replace(replace(request.querystring("New"),"#","/"),"../","")
         sOpenURL = request.querystring("Open")
			sNewURL = request.querystring("New")
			sFName = request.querystring("fname")

	'MOVE THE FILE 
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	Image_Folder = fn_get_sites_folder(Store_Id,"Images")
	
	if (request.querystring("fname")) <> "" then
	Session("move_path") =  Image_Folder&sOpenPath&"\"&sFName
	end if
	On Error Resume Next

 	Source_Path = Session("move_path")
	Dest_Path = Image_Folder&sNewPath&"\"&sFName


	if (request.querystring("act")) = "M" then
		dim FileObject
		set FileObject=Server.CreateObject("Scripting.FileSystemObject")
		FileObject.MoveFile Source_Path,Dest_Path
		set FileObject=nothing
				%>
				<SCRIPT LANGUAGE="JavaScript">

function refreshParent() {
window.opener.location.href = window.opener.location.href;

if (window.opener.progressWindow)

{
window.opener.progressWindow.close()
}
window.close();
}

refreshParent();
				</SCRIPT>
				<%

	end if

	if (request.querystring("act")) = "S" then
			sOpenPath = ""
	end if

%>

	<td width="100" height="50" align="left" valign=top>
     <%
			strHTML = strHTML & ("<table width='100%' class=list>")


				folderspec = Image_Folder&"\"&sOpenPath

				Dim fso, f, f1, fc
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.GetFolder(folderspec)
				Set fc = f.Files
				Set ff = f.SubFolders
				tfiles = 0
				str_class=1

				      strHTML = strHTML & ("<tr class="&str_class&">")
						strHTML = strHTML & ("<td><a class='link' href='move_image.asp?New="&server.urlencode(sNewURL)&"&Open="&server.urlencode(sOpenURL)&"&"&sAddString&"'><img src=images/folder.gif border=0></a></td><td>Top Level</td>")
			         strHTML = strHTML & ("<td ><a class='link' href='move_image.asp?act=M&New="&server.urlencode(sNewURL)&"&Open="&server.urlencode(sOpenUrl)&"&"&sAddString&"'>Select</a> </td>")

				For Each f1 in ff
					  if str_class=1 then
					  str_class=0
						 else
						str_class=1
						 end if
						strHTML = strHTML & ("<tr class="&str_class&">")
						strHTML = strHTML & ("<td><a class='link' href='move_image.asp?New="&server.urlencode(sNewURL&"#" & f1.name)&"&Open="&server.urlencode(sOpenURL)&"&"&sAddString&"'><img src=images/folder.gif border=0></a></td><td>"&f1.name&"</td>")

			      strHTML = strHTML & ("<td ><a class='link' href='move_image.asp?act=M&New="&server.urlencode(sNewURL&"#" & f1.name)&"&Open="&server.urlencode(sOpenUrl)&"&"&sAddString&"'>Select</a> </td>")

            tfiles = tfiles + 1
				Next

		strHTML = strHTML & ("</table>") %>

		<table border=0 cellpadding=0 cellspacing=0  D="Table4">
				<tr><td>
				<div class="bar" style="padding-left: 5px;">
				<font size="2" face="tahoma" color="black"><b>Folders</b></font>
				</div>
				</td></tr>
				</table>
				
				<%=strHTML%>
				

		</td>
