<!--#include file="global_settings.asp"-->
<html>
<head>
	<meta name="Pragma" content="no-cache">
	<META HTTP-EQUIV="pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Easystorecreator</title>

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

      <table>
	
	<tr>
	<td width="250" colspan="3" height="74" align="center">
			<table width="240">
				
				<tr>
					<td><b>Name</b></td>

					<td align="center">&nbsp;</td>

				</tr>
		
				<%
				Download_Folder = fn_get_sites_folder(Store_Id,"Download")
				Dim fso, f, f1, fc
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.GetFolder(Download_Folder)
				Set fc = f.Files
				tfiles = 0
				str_class=1
				For Each f1 in fc %>
				      <% if str_class=1 then
					  str_class=0
				       else
					   str_class=1
				       end if %>
					<tr class="<%= str_class %>">
						<td><%= f1.name %></td>
						<td align="center"><a class="link" href="JavaScript:setParentResults('<%= request.queryString("returnArg") %>','<%= f1.name %>');">Select</a> </td>

				<% tfiles = tfiles + 1
				Next %>
		</table>
		</td>
	</tr>
	
	<tr>
	<td colspan="3" height="25" valign="top">&nbsp;</td>
	</tr>
	</table>
	</body>
	</html>

