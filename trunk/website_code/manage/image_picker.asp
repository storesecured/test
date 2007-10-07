<!--#include file="Global_Settings.asp"-->
<html>
<head>
	<meta name="Pragma" content="no-cache">
	<META HTTP-EQUIV="pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Image Picker</title>

<SCRIPT LANGUAGE = "JavaScript">

	<!--
	function setParentResults(resArg,resVal) {
		window.parent.opener.setResults(resArg, resVal);
		window.parent.close();}
	//-->

</SCRIPT>
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

sOpenPath = replace(replace(request.querystring("Open"),"#","/"),"../","")
sOpenURL = request.querystring("Open")

%>
<SCRIPT LANGUAGE = "JavaScript">
	
	<!--
	function setParentResults(resArg,resVal) {
		if (resArg == 'HTML') {
			window.opener.idContent.InsertImage('<%= Site_Name %>/images/<%= sOpenPath %>/'+resVal,"", "", "","" ,"","","");
			window.parent.close();
		} else {
		    <% if sOpenPath = "" then %>
		    window.parent.opener.setResults(resArg, resVal);
		    <% else %>
		    <% 'sOpenPath2 = Mid(sOpenPath,2,len(sOpenPath)-2) %>
		    window.parent.opener.setResults(resArg, '<%= sOpenPath %>/'+resVal);
		    <% end if %>
			window.parent.close();
		}
	}
	//-->
	
	function previewImage(sImage) {
		document.images['PRV_IMG'].src='<%=Site_Name %>images<%= sOpenPath %>/'+sImage;
	}

</SCRIPT>
	<table valign=top>
	<tr>
	<td width="200" height="74" align="center" valign=top>
      <%
			strHTML = strHTML & ("<table width='100%' class='list'>")

				strHTML = strHTML & ("<tr><td><b>Name</b></td>")
				strHTML = strHTML & ("<td><b>Preview</b></td>")
				strHTML = strHTML & ("<td><b></b></td></tr>")

                on error resume next
				Image_Folder = fn_get_sites_folder(Store_Id,"Images")&sOpenPath
            
				Dim fso, f, f1, fc
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.GetFolder(Image_Folder)
				Set fc = f.Files
				Set ff = f.SubFolders
				tfiles = 0
				str_class=1
				if sOpenPath <> "" then
				strHTML = strHTML & ("<tr><td><a class='link' href='left_image_picker.asp'><img src=images/folderopen.gif border=0></a></td><td><a class='link' href='image_picker.asp?returnArg="&request.querystring("returnArg")&"'>Top Level</a></td></tr>")

				end if
				For Each f1 in ff
					  if str_class=1 then
					  str_class=0
						 else
						str_class=1
						 end if
						strHTML = strHTML & ("<tr class="&str_class&">")
						strHTML = strHTML & ("<td><a class='link' href='image_picker.asp?Open="&server.urlencode(sOpenURL&"#" & f1.name)&"&returnArg="&request.querystring("returnArg")&"'><img src=images/folder.gif border=0></a></td><td>"&f1.name&"</td>")

						strHTML = strHTML & ("<td align='center'><a class='link' href='image_picker.asp?Open="&server.urlencode(sOpenURL&"#" & f1.name)&"&returnArg="&request.querystring("returnArg")&"'>Open</a> </td>")
			      strHTML = strHTML & ("<td>&nbsp;</td>")
           tfiles = tfiles + 1
				Next
				
				i=0
            sStart=0
            if request.querystring("start") <>"" then
               sStart=int(request.querystring("start"))
            end if
            sEnd=sStart+250

				For Each f1 in fc
               if i>=sStart and i<sEnd then
				      if str_class=1 then
					      str_class=0
				      else
					      str_class=1
				      end if
                  sFileName = replace(f1.name," ","%20")
					   strHTML = strHTML & ("<tr class='"&str_class&"'>")
						strHTML = strHTML & ("<td><img src=images/file.gif></td><td>"&lcase(left(f1.name,25))&"</td>")
						strHTML = strHTML & ("<td><a class='link' href=""JavaScript:previewImage('"&f1.name&"');"">Preview</a></td>")
						strHTML = strHTML & ("<td align='center'><a class='link' href=""JavaScript:setParentResults('"&request.queryString("returnArg")&"','"&sFileName&"');"">Select</a> </td>")

				      tfiles = tfiles + 1
              end if
              i=i+1
				Next
				
				sNext=0
            sPrev=0
            if sStart>0 then
               sPrev=1
            end if
            if i>sEnd then
               sNext=1
            end if

	       strHTML = strHTML & ("</table>") %>
		<table border=0 cellpadding=0 cellspacing=0 width=310 ID="Table4">
				<tr><td>
				<div class="bar" style="padding-left: 5px;">
				<font size="2" face="tahoma" color="black"><b>File Name</b></font>
				</div>
				</td></tr>
				</table>
				
				<div style="overflow:auto;height:377;width:370;BORDER-LEFT: #316AC5 1px solid;BORDER-RIGHT: LightSteelblue 1px solid;BORDER-BOTTOM: LightSteelblue 1px solid;">
				<%=strHTML%>
				</div>

		</td>
		<td width="10%"></td>
		<td align=left valign=top width="90%"><table border="1" cellspacing="0" cellpadding=2 class="list"><tr><td><b>Image Preview</b><BR>
	<img src=images/spacer.gif name="PRV_IMG"></td></tr></table></td>
	</tr>
	<tr><td><table>
	<% if sPrev=1 then %>
        <TR><td colspan=2 align=right><a href=image_picker.asp?Open=<%= server.urlencode(sOpenURL)%>&Start=<%= sStart-250%>&returnArg=<%= request.querystring("returnArg") %> class=link>Prev</a> | </td>
   <% else %>
        <td colspan=2 align=right>&nbsp; | </td>
   <% end if %>
   <% if sNext=1 then %>
        <td colspan=2 align=left><a href=image_picker.asp?Open=<%= server.urlencode(sOpenURL)%>&Start=<%= sEnd%>&returnArg=<%= request.querystring("returnArg") %> class=link>Next</a></td>
   <% else %>
        <td colspan=2 align=left>&nbsp;</td></tr>
   <% end if %>
	</table></td></tr></table>
	</body>
	</html>

