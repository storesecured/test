<!--#include file="global_settings.asp"-->
<!--#include file="include/UploadScript.asp"-->

<%
server.scripttimeout=180
assetFolder =  fn_get_sites_folder(Store_Id,"Root")
imgFolder =  fn_get_sites_folder(Store_Id,"Root")

function replaceSpecial (sName)
   sNewName = sName
   sNewName = replace(sNewName," ","_")
   sNewName = replace(sNewName,"<","")
   sNewName = replace(sNewName,">","")
   sNewName = replace(sNewName,"/","")
   sNewName = replace(sNewName,"?","")
   sNewName = replace(sNewName,"&","")
   replaceSpecial = sNewName
end function

dim objFSO
dim objMainFolder

dim strOptions



dim imgobjMainFolder 'main images directory of the store
dim strimgOptions  'for displaying the directory options for the files

dim strHTML
dim catid


dim filecat			'var used for show/hide of the files dropdown
filecat=""

dim deptcat			'var used for show/hide of the dept dropdown
deptcat=""

dim fileid			'var to catch the file directory option selected 

strHTML = ""

set objFSO = server.CreateObject ("Scripting.FileSystemObject")
set objMainFolder = objFSO.GetFolder(assetFolder)
set imgobjMainFolder = objFSO.GetFolder(imgFolder)

fileid = CStr(request("fileid"))
deptid = CStr(request("deptid"))
catid = CStr(request("catid"))'bisa form, bisa querystring
if catid="" then catid = "Departments"

dim objTempFSO
dim objTempFolder
dim objTempFiles
dim objTempFile

set objTempFSO = server.CreateObject ("Scripting.FileSystemObject")

strHTML = strHTML & "<table border=0 cellpadding=3 cellspacing=0 width=360>"

if catid="Departments" then
    set deptfields1=server.createobject("scripting.dictionary")
    sql_select="SELECT top 1000 full_name, department_name, department_id FROM Store_dept WITH (NOLOCK) WHERE Store_id="&Store_Id&" order by Full_Name"
    Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
    i=0
    FOR deptrowcounter1= 0 TO deptfields1("rowcount")
        Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
        Department_Name = deptdata1(deptfields1("department_name"),deptrowcounter1)
        sub_department_id = deptdata1(deptfields1("department_id"),deptrowcounter1)
        sLink = fn_dept_url(Full_Name,"")
        strHTML = strHTML & "<tr bgcolor=Lavender>"
        strHTML = strHTML & "<td valign=top>" & Full_Name & "</td>"
        strHTML = strHTML & "<td valign=top>Id:" & Sub_Department_ID & "</td>"
        '  	 strHTML = strHTML & "<td valign=top style=""cursor:hand;"" onclick=""selectFile('" & sLink & "')""><u><font color=blue>select</font></u></td></tr>"
        strHTML = strHTML & "<td valign=top style=""cursor:hand;"" onclick=""selectFile('" & sLink & "');applyLink();self.close();""><u><font color=blue>select</font></u></td></tr>"
         i=i+1
    Next
elseif catid="Items" then
  set deptfields1=server.createobject("scripting.dictionary")
  sql_select="exec wsp_items_display_all "&Store_Id&","&deptid&";"
  Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
  i=0
  FOR deptrowcounter1= 0 TO deptfields1("rowcount")
     Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
  	 item_name = deptdata1(deptfields1("item_name"),deptrowcounter1)
  	 item_page_name = deptdata1(deptfields1("item_page_name"),deptrowcounter1)
  	 item_id = deptdata1(deptfields1("item_id"),deptrowcounter1)
  	 sLink = fn_item_url (full_name,item_page_name)
  	 strHTML = strHTML & "<tr bgcolor=Lavender>"
     strHTML = strHTML & "<td valign=top>" & Item_Name & "</td>"
  	 strHTML = strHTML & "<td valign=top>Id:" & Item_id & "</td>"
  	 strHTML = strHTML & "<td valign=top style=""cursor:hand;"" onclick=""selectFile('" & sLink & "');applyLink();self.close();""><u><font color=blue>select</font></u></td></tr>"
        i=i+1
  Next
elseif catid="Pages" then
  set deptfields1=server.createobject("scripting.dictionary")
  sql_select="SELECT top 1000 Page_Id, Page_Name, File_Name from store_pages WITH (NOLOCK) where Store_id="&Store_Id&" order by page_name"
  Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
  i=0
  FOR deptrowcounter1= 0 TO deptfields1("rowcount")
  	 Page_Id = deptdata1(deptfields1("page_id"),deptrowcounter1)
  	 Page_Name = deptdata1(deptfields1("page_name"),deptrowcounter1)
  	 File_Name = deptdata1(deptfields1("file_name"),deptrowcounter1)
	 sLink = fn_page_url(File_name,Is_Link)
  	 strHTML = strHTML & "<tr bgcolor=Lavender>"
     strHTML = strHTML & "<td valign=top>" & Page_Name & "</td>"
  	 strHTML = strHTML & "<td valign=top>Id:" & Page_Id & "</td>"
  	 strHTML = strHTML & "<td valign=top style=""cursor:hand;"" onclick=""selectFile('" & sLink & "');applyLink();self.close();""><u><font color=blue>select</font></u></td></tr>"
         i=i+1
  Next
elseif catid="Files" then

	set objTempFolder = objTempFSO.GetFolder(fileid)
	set objTempFiles = objTempFolder.files

	for each objTempFile in objTempFiles
 'if instr(objTempFile.type,"Image") > 0 then
 if i<3000 then
  	'***********
  	'objTempFile.path => image physical path
  	'basePath => base path
  	set basePath = objFSO.GetFolder(imgFolder)
  	PhysicalPathWithoutBase = Replace(objTempFile.path,basePath.path,"")	
  	sTmp = replace(PhysicalPathWithoutBase,"\","/")'fix from physical to virtual
  	sCurrImgPath = imgFolder & sTmp

  	'***********
  	
  	strHTML = strHTML & "<tr bgcolor=Lavender>"
  	strHTML = strHTML & "<td valign=top>" & objTempFile.name & "</td>"
  	strHTML = strHTML & "<td valign=top>" & objTempFile.type & "</td>"
  	'strHTML = strHTML & "<td valign=top>" & FormatNumber(objTempFile.size/1000,0) & " kb</td>"

		strHTML = strHTML & "<td valign=top style=""cursor:hand;"" onclick=""selectFile('" & sLink & "');applyLink();self.close();""><u><font color=blue>select</font></u></td>"

   
 'end if
      i=i+1
 end if
next

end if
strHTML = strHTML & "</table>"
if i>=1000 then
   strHTML = strHTML & "Your total number of lines exceeds 1000.  This list will only show the first 1000 of each type."
end if

Function createDepartmentDropdown(this_department_id)
    stringOptions=""
    on error goto 0
    set deptfields1=server.createobject("scripting.dictionary")
    sql_select="SELECT top 1000 full_name, department_id FROM Store_dept WITH (NOLOCK) WHERE Store_id="&Store_Id&" order by Full_Name"
    Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
    FOR deptrowcounter1= 0 TO deptfields1("rowcount")
        Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
        sub_department_id = deptdata1(deptfields1("department_id"),deptrowcounter1)
        if cstr(this_department_id)=cstr(sub_department_id) then
            sSelected="selected"
        else
            sSelected=""
        end if
        stringOptions = stringOptions & "<option value=""" & sub_department_id & """ "&sSelected&">" & Full_Name	 & "</option>" & vbCrLf
    Next
    createDepartmentDropdown=stringOptions

end function

'for files directory dropdown
Function createImgCategoryOptions(pi_objFolder) 
    dim objFolder
    dim objFolders
	
    set objFolders = pi_objfolder.SubFolders
    for each objFolder in objFolders 
		'Recursive programming starts here
		createImgCategoryOptions objFolder
    next

    if pi_objFolder.attributes and 2 then
		'hidden folder then do nothing
	elseif instr(pi_objFolder.path,"aspnet_client")>0 then
	    'aspnet system files, do nothing
	else	

'***********
set basePath = objFSO.GetFolder(server.MapPath(imgFolder))
'response.Write catid & " - " & oo.path
Response.Write Replace(pi_objFolder.path,basePath.path,"")	
'***********
	
		
		if CStr(fileid)=CStr(pi_objFolder.path) then
			strimgOptions = strimgOptions & "<option value=""" & pi_objFolder.path & """ selected>" & Replace(pi_objFolder.path,basePath.path,"")	 & "</option>" & vbCrLf

		
		else
			strimgOptions = strimgOptions & "<option value=""" & pi_objFolder.path & """>" & Replace(pi_objFolder.path,basePath.path,"")	  & "</option>" & vbCrLf
		end if
    end if
    
    strimgOptions = strimgOptions & vbCrLf
    createImgCategoryOptions = strimgOptions
End Function



Function createCategoryOptions(catid)

    if catid="Departments" then
       deptselected ="selected"
	   filecat="none"
	   deptcat="none"
    elseif catid="Items" then
       itemselected ="selected"
   	   filecat="none"
   	   deptcat="block"
    elseif catid="Pages" then
       pageselected="selected"
  	   filecat="none"
  	   deptcat="none"
   elseif catid="Files" then
		fileselected="selected"
	   filecat="block"
	   deptcat="none"
		
    end if
    strOptions = strOptions & "<option value=""Departments"" "&deptselected&">Departments</option>" & vbCrLf
    strOptions = strOptions & "<option value=""Items"" "&itemselected&">Items</option>" & vbCrLf
    strOptions = strOptions & "<option value=""Pages"" "&pageselected&">Pages</option>" & vbCrLf
    strOptions = strOptions & "<option value=""Files"" "&fileselected&">Files</option>" & vbCrLf
    strOptions = strOptions & vbCrLf
    createCategoryOptions = strOptions
End Function
%>
<html>
<head>
	<title>Store Files</title>
	
	<style>

	BODY
		{
		FONT-FAMILY: Verdana;FONT-SIZE: xx-small;
		}
	TABLE
		{
	    FONT-SIZE: xx-small;
	    FONT-FAMILY: Tahoma
		}
	INPUT
		{
		font:8pt verdana,arial,sans-serif;
		}
	select
		{
		height: 22px; 
		top:2;
		font:8pt verdana,arial,sans-serif
		}	
	.bar 
		{
		BORDER-TOP: #99ccff 1px solid; BACKGROUND: #336699; WIDTH: 100%; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 20px
		}		
	</style>
		
<script language="Javascript" src="editor/scripts/innovaeditor.js"></script>

	<script language="JavaScript">

	function selectFile(sURL)
		{
			document.getElementById("panel_Flash").display = "none";
			document.getElementById("panel_Link").display = "none";

			document.getElementById("inpSRC").value = sURL;
		var arrTmp = document.getElementById("inpSRC").value.split(".");
		var FileExtension = arrTmp[arrTmp.length-1]
		var MediaType = ""
		if(FileExtension=="wmv"||FileExtension=="avi") MediaType = "Video"
		if(FileExtension=="wma"||FileExtension=="wav"||FileExtension=="mid") MediaType = "Sound"
		if(FileExtension=="gif"||FileExtension=="jpg") MediaType = "Image"
		if(FileExtension=="swf") MediaType = "Flash"
				
	
		} 
	function applyLink()
		{


				if(navigator.appName.indexOf('Microsoft')!=-1){/*For IE browser, use dialogArguments.oUtil.obj 
					to get the active editor object*/
//					var oEdit=dialogArguments.oUtil.obj;
						var oEdit=window.opener.oUtil.obj;
					oEdit.insertLink(inpSRC.value,inpLinkText.value);
				}	else {/*For Mozilla browsers, use window.opener.oUtil.obj 
					to get the active editor object*/
					var oEdit=window.opener.oUtil.obj;
					oEdit.insertLink(inpSRC.value,inpLinkText.value);
				}
		}
	</script>
</head>
<body>
		<table border=0 cellpadding=3 cellspacing=3 ID="Table2">
		<form method=post action=editor_internal.asp id=form2 name=form2>
        <tr>
  		<td valign=top>
    
						<table border=0 height=30 cellpadding=0 cellspacing=0 ID="Table3"><tr>
						<td><b>Select Type&nbsp;:&nbsp;</b></td>
						<td>
						<select id=catid name=catid onchange="form2.submit()">
						<%=createCategoryOptions(catid)%>
						</select> 
						</td></tr></table>

						<div id=filecat style="display:<%=filecat%>">
						<table border=0 height=30 cellpadding=0 cellspacing=0 ID="Table3"><tr>
						<td><b>Select Directory&nbsp;:&nbsp;</b></td>
						<td>
						<select id=fileid name=fileid onchange="form2.submit()">
						<%=createImgCategoryOptions(imgObjMainFolder)%>				
						</select>
						</td></tr>
						</table>
						</div>
						
						<div id=deptcat style="display:<%=deptcat%>">
						<table border=0 height=30 cellpadding=0 cellspacing=0 ID="Table5"><tr>
						<td><b>Select Department&nbsp;:&nbsp;</b></td>
						<td>
						<select id=deptid name=deptid onchange="form2.submit()">
						<%=createDepartmentDropdown(deptid)%>				
						</select>
						</td></tr>
						</table>
						</div>
				
				
				<table border=0 cellpadding=0 cellspacing=0 width=358 ID="Table4">
				<tr><td>
				<div class="bar" style="padding-left: 5px;">
				<font size="2" face="tahoma" color="white"><b><%= catid %> Name</b></font>
				</div>
				</td></tr>
				</table>
				
				<div style="overflow:auto;height:260;width:360;BORDER-LEFT: #316AC5 1px solid;BORDER-RIGHT: LightSteelblue 1px solid;BORDER-BOTTOM: LightSteelblue 1px solid;">
				<%=strHTML%>
				</div>

		</td>						
		</tr>
		</form>
		</table>

<table cellpadding=3 cellspacing=0 ID="Table1">

<!-- Source -->
<INPUT type="hidden" ID="inpSRC" NAME="inpSRC" size=30>
<!-- /Source -->

<!-- Media Type -->
<!--
<tr>
<td valign=top>Media Type: </td><td>
	<INPUT type="text" ID="inpMediaType" NAME="inpMediaType" size=7 disabled>
	</td>
</tr>
-->
<!-- /Media Type -->

<!-- Insert_Link -->
<tr id="panel_Link" style="display:none">
<td>Link Text: </td><td><INPUT type="text" ID="inpLinkText" NAME="inpLinkText"></td>
</tr>
<!-- /Insert_Link -->

<!-- Insert_Flash -->
<tr id="panel_Flash" style="display:none">
<td>
	Dimension:  
</td>
<td>
	Width: <INPUT type="text" ID="inpFlashWidth" NAME="inpFlashWidth" size=4 value=150>
	Height: <INPUT type="text" ID="inpFlashHeight" NAME="inpFlashHeight" size=4 value=150>
</td>
</tr>
<!-- /Insert_Flash -->



</table>
<br>

<INPUT type="button" value="CANCEL" onclick="self.close();" ID="Button2" NAME="Button2">

</body>
</html> 

