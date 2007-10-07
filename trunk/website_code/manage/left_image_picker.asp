<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
Server.ScriptTimeout = 4800

	If Form_Error_Handler(Request.Form) <> "" then 
    Error_Log = Form_Error_Handler(Request.Form)
	%> <!--#include file="Include/Error_Template.asp"--><%
	Response.end
end if


'Paging yet to be done
On Error Resume Next




sOpenPath = replace(replace(request.querystring("Open"),"#","\"),"..\","")
sOpenURL = request.querystring("Open")
Image_Folder = fn_get_sites_folder(Store_Id,"Images")
Session("move_path") =  Image_Folder&sOpenPath

'CODE FOR DELETING AN IMAGE
Image_Del = Request.QueryString("Image_Del")

if Image_Del <> "" then
	'DELETE THE FILE FROM DISK
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	File_Full_Name = Image_Folder&sOpenPath&"\"&Image_Del

	FileObject.DeleteFile(File_Full_Name)
	Set FileObject = Nothing
end if

Folder_Del = Request.QueryString("Folder_Del")
if Folder_Del <> "" then
	'DELETE THE FILE FROM DISK
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	File_Folder = Image_Folder&sOpenPath&"\"&Folder_Del
	On Error Resume Next
	FileObject.DeleteFolder(File_Folder)
	Set FileObject = Nothing
end if

if  Request.Form("Delete_Item") <> "" then
	if Request.Form("Delete_Item")="SEL" then
	  if request.form("DELETE_IDS") <> "" then

		  Set FileObject = CreateObject("Scripting.FileSystemObject")
		  for each this_file_name in Split(request.form("DELETE_IDS"), ",")
			this_file_name = trim(this_file_name)

			if this_file_name <> "" then
				File_Full_Name = Image_Folder&sOpenPath&"\"&this_file_name
				On Error resume next
				FileObject.DeleteFile(File_Full_Name)
			end if
		  next
		  Set FileObject = Nothing
		end if
	end if
end if

on error resume next
if request.form("NewFolder") <> "" or Request.form("Folder_name") <> "" then
This_Folder = Image_Folder&sOpenPath&""

	on error goto 0
  Set FileObject = CreateObject("Scripting.FileSystemObject")
	Folder_Name= request.form("Folder_Name")
	This_Folder = This_Folder & "\" & Folder_Name

	If not FileObject.FolderExists(This_Folder) Then
			set f1 = FileObject.CreateFolder(This_Folder)
	end if
	Set FileObject = Nothing
	on error resume next
end if
sTextHelp="images/filesimages.doc"

if request.form("ResizeFolder") <> "" then
  Folder_Path = Image_Folder&sOpenPath&"\"
  sMaxWidth = 150
  sMaxHeight = 150
  if request.form("max_width") <> "" then
     sMaxWidth=request.form("max_width")
  end if
  if request.form("max_height") <> "" then
     sMaxHeight=request.form("max_height")
  end if
  %><!--#include file="resize2screen.asp"--><%
end if

sFormAction = "left_image_picker.asp?Open="&server.urlencode(sOpenURL)&"&Start="&request.querystring("Start")
sTitle = "List/Preview Files/Images"
sFullTitle = "Design > Files/Images"
thisRedirect = "left_image_picker.asp"
sMenu = "design"
sQuestion_Path = "images/image_list.htm"
createHead thisRedirect

%>


<SCRIPT LANGUAGE = "JavaScript">
	
	<!--
	function setParentResults(resArg,resVal) {
		window.parent.opener.setResults(resArg, resVal);
		window.parent.close();}
	//-->
	
	function previewImage(sImage) {
		document.images['PRV_IMG'].src='<%=Site_Name %>images<%= replace(sOpenPath,"\","/") %>/'+sImage;
    }

	<!--
	function VerifyDelete(OID)
	{
		return confirm("Are you sure you want to delete this order?");}
	// -->

	  function goDeleteSeleted(){
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked) 
				tsel = tsel+1;

		if (confirm("Are you sure you want to delete these images?")){
			document.forms[0].Delete_Item.value="SEL";
			document.forms[0].submit(); }}

		function selectAllItems(){
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
				document.forms[0].DELETE_IDS[i].checked = document.forms[0].selectall.checked; }



//function for deleting folders / files (with alert message)
function ConfirmDelete(stype, objName)
{
	if (stype.toUpperCase() =="FOLDER")
	{
		if (confirm("Are you sure you want to delete this Folder?"))
			document.location.href="left_image_picker.asp?Folder_Del="+objName+"&Open=<%=server.urlencode(sOpenURL)%>&<%=sAddString%>"
	}
	else if(stype.toUpperCase() == "FILE")
	{
		if (confirm("Are you sure you want to delete this File?"))			
	 document.location.href="left_image_picker.asp?Image_Del="+escape(objName)+"&Open=<%=server.urlencode(sOpenURL)%>&<%=sAddString%>"
	}
}

//function for moving images to folders  (with alert message)
function goMoveImage(objName)
{
		var w = 400;
		var h = 500;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="move_image.asp?act=S&fname="+escape(objName)+"&Open=<%= server.urlencode(request.querystring("Open")) %>&<%=sAddString%>"
		reWin = window.open(theUrl,'moveImage','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,'+winprops);
}

// Valid characters for Folder Name
function isAllCharacters(objValue)
	{

		var characters="1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
		var tmp
		ltag = 0
		temp = (objValue.length)
		for (var i=0;i<=temp;i++)
		{
			tmp=objValue.substring(i,i+1)
			if (characters.indexOf(tmp)==-1)
			{
				ltag = 1
			}
		}
		if(ltag == 1)
			return false
		else
			return true
	}
	
// Function to submit Add Folder
function fnSubmit()
{
	var fName = document.forms[0].Folder_Name.value
	if (isAllCharacters(fName) == false)
	{
	alert("Folder name should not contain any special characters or blank space.")
	document.forms[0].Folder_Name.focus();
	return false;
	}
	else
	{
	document.forms[0].submit()
	}
}
</SCRIPT>
<input type="Hidden" name="Delete_Item" value="">

	<TR bgcolor='#FFFFFF'>

		<td colspan="4" height="23">
			<input type="button" class="Buttons" value="Upload New Images" name="Add" OnClick='JavaScript:self.location="upload_images.asp?Open=<%= server.urlencode(sOpenURL) %>&<%= sAddString %>"'></td></tr>

        <TR bgcolor='#FFFFFF'>

		<td colspan="4" height="23">Folder Name <input type="text" name="Folder_Name" value="<%=Folder_Name%>"><input type="button" class="Buttons" value="Add Folder" name="NewFolder" onclick='javascript:return fnSubmit(2)'><br><!--onclick='javascript:return fnSubmit(2)'-->
        <input type="hidden" name="Folder_Name_C" value="|Specials|0|100|||Folder Name">

		</td>
	</tr>
	
	<TR bgcolor='#FFFFFF'>
	<td width="200" height="74" align="center" valign=top>
     <%
			strHTML = strHTML & ("<table width='100%' class=list>")


				This_Folder = Image_Folder&"\"&sOpenPath

				Dim fso, f, f1, fc
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.GetFolder(This_Folder)
				Set fc = f.Files
				Set ff = f.SubFolders

iRecCnt = f.Files.Count

Dim theFiles( )
ReDim theFiles( 500 ) ' arbitrary size!
currentSlot = -1 ' start before first slot
i=0
sStart=0
if request.querystring("start") <>"" then
   sStart=int(request.querystring("start"))
end if
sEnd=sStart+50

For Each f1 in fc
  if i>=sStart and i<sEnd then
    filename = f1.Name
    ftype = f1.Type
    currentSlot = currentSlot + 1
    If currentSlot > UBound( theFiles ) Then
    		ReDim Preserve theFiles( currentSlot + 99 )
    End If
    '  put in an array!
    theFiles(currentSlot) = Array(filename,ftype)
    'theFiles(currentSlot) = filename
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

fileCount = currentSlot 
ReDim Preserve theFiles( currentSlot ) 

				tfiles = 0
				str_class=1
				if sOpenPath <> "" then
				strHTML = strHTML & ("<TR bgcolor='#FFFFFF'><td><a class='link' href='left_image_picker.asp'><img src=images/folderopen.gif border=0></a></td><td><a class='link' href='left_image_picker.asp'>Top Level</a></td></tr>")

				end if
				For Each f1 in ff
					  if str_class=1 then
					  str_class=0
						 else
						str_class=1
						 end if
						strHTML = strHTML & ("<tr class="&str_class&">")
						strHTML = strHTML & ("<td><a class='link' href='left_image_picker.asp?Open="&server.urlencode(sOpenURL&"#" & f1.name)&"&"&sAddString&"'><img src=images/folder.gif border=0></a></td><td>"&f1.name&"</td>")

						strHTML = strHTML & ("<td align='center'><a class='link' href='left_image_picker.asp?Open="&server.urlencode(sOpenURL&"#" & f1.name)&"&"&sAddString&"'>Open</a> </td>")
			      strHTML = strHTML & ("<td align='center'><a class='link' href=""JavaScript:ConfirmDelete('FOLDER','"&f1.name&"');"">Del</a> </td>")

            strHTML = strHTML & ("<td>&nbsp;</td>")
            tfiles = tfiles + 1
				Next

			

			iRecCnt = f.Files.Count



			For i = 0 To fileCount
                                       if theFiles(i)(0)<>"spacer.gif" then

  					  if str_class=1 then
  					  str_class=0
  						 else
  						str_class=1
  						 end if
  
  					   strHTML = strHTML & ("<tr class='"&str_class&"'>")
                                                if instr(theFiles(i)(1),"Image") > 0 then
  						strHTML = strHTML & ("<td><a class='link' href=""JavaScript:previewImage('thmbnl_"&theFiles(i)(0)&"');""><img src=images/file.gif border=0></a></td><td>"&lcase(left(theFiles(i)(0),25))&"</td>") & vbCrlf
  						strHTML = strHTML & ("<td><a class='link' href=""JavaScript:previewImage('"&theFiles(i)(0)&"');"">View</a>") & vbCrlf
  						else
  						strHTML = strHTML & ("<td><img src=images/file.gif></td><td>"&theFiles(i)(0)&"</td><td>") & vbCrlf
              end if
    						strHTML = strHTML & ("</td><td align='center'><a class='link' href=""JavaScript:ConfirmDelete('FILE','"&theFiles(i)(0)&"');"">Del</a> </td>") & vbCrlf
  						strHTML = strHTML & ("</td><td align='center'><a class='link' href=""Javascript:goMoveImage('"&theFiles(i)(0)&"');"">Move</a> </td>") & vbCrlf
  						strHTML = strHTML & ("<td align=center><input class='image' type='checkbox' name='DELETE_IDS' value='"&theFiles(i)(0)&"'>") & vbCrlf
  				end if
                                  tfiles = tfiles + 1

				Next
				if tfiles > 1 then
  				   strHTML = strHTML & ("<TR bgcolor='#FFFFFF'><td colspan=4></td><td align='center'>")
  				   strHTML = strHTML & ("<input class='image' type='checkbox' name='selectall' onClick=""JavaScript:selectAllItems();""><br>Check All</td></tr>")
  				   strHTML = strHTML & ("<TR bgcolor='#FFFFFF'><td colspan=2 align=left><input type='Button' name='DeleteAll' value='Delete Selected' onClick=""JavaScript:goDeleteSeleted();""></td>")
  				   if sPrev=1 then
                 strHTML = strHTML & ("<td colspan=2 align=right><a href=left_image_picker.asp?Open="&server.urlencode(sOpenURL)&"&Start="&sStart-50&" class=link>Prev</a></td>")
               End if
               if sNext=1 and sPrev=1 then
                 strHTML = strHTML & ("<td colspan=2 align=center> | </td>")
               
               end if
			if sNext=1 then
                 strHTML = strHTML & ("<td colspan=2 align=left><a href=left_image_picker.asp?Open="&server.urlencode(sOpenURL)&"&Start="&sEnd&" class=link>Next</a></td>")
               else
                 strHTML = strHTML & ("<td colspan=2 align=left>&nbsp;</td></tr>")
               end if
				end if
		strHTML = strHTML & ("</table>") %>

		<table border=0 cellpadding=0 cellspacing=0 width=310 ID="Table4">
				<TR bgcolor='#FFFFFF'><td>
				<div class="bar" style="padding-left: 5px;">
				<font size="2" face="tahoma" color="black"><b>File Name</b></font>
				</div>
				</td></tr>
				</table>
				
				<div style="overflow:auto;height:320;width:377;BORDER-LEFT: #316AC5 1px solid;BORDER-RIGHT: LightSteelblue 1px solid;BORDER-BOTTOM: LightSteelblue 1px solid;">
				<%=strHTML%>
				</div>

		</td>
		<td align=center valign=top width="90%" rowspan=2><b>Image Preview</b><BR>
	<img src="images/spacer.gif" name="PRV_IMG"></td></tr>

	<TR bgcolor='#FFFFFF'><td><BR>
  <table class=list>
  <TR bgcolor='#FFFFFF'><td colspan=2>Create thumbnails for currently shown images in this folder</td></tr>
  <TR bgcolor='#FFFFFF'><td>Max Width</td><td><input type=text name=max_width value=150 size=5>pixels</td></tr>
  <TR bgcolor='#FFFFFF'><td>Max Height</td><td><input type=text name=max_height value=150 size=5>pixels</td></tr>
  <TR bgcolor='#FFFFFF'><td colspan=2><input type="submit" class="Buttons" value="Create Thumbnails" name="ResizeFolder"></td></tr></table>
  </td></tr>

 

<% createFoot thisRedirect, 0%>
