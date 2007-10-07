<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
'Stores only files with size less than MaxFileSize
sOnlyFiles = "all"

Destination_Folder = fn_get_sites_folder(Store_Id,"Download")
%>
<!--#include file="include/huge_upload.asp"-->
<%


sFormAction = PostURL
sName = "upload_file"
sTitle = "ESD File Upload"
sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=file_list.asp class=white>ESD Files</a> > Upload"
sSubmitName = "Upload_File"
thisRedirect = "upload_files.asp"
sEncType = "ENCTYPE='multipart/form-data'"
sMenu = "inventory"
sQuestion_Path = "inventory/upload_esd_files.htm"
createHead thisRedirect

if Service_Type < 7 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>

<%
elseif Trial_Version then %>
  <TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available for trial stores.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% else
%>
</form><form method="POST" action="<%=PostURL %>" ENCTYPE="multipart/form-data" OnSubmit="return ProgressBar();" >

	<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File1" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File2" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File3" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File4" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File5" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File6" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File7" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File8" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File9" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="File10" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%" colspan=3 align=center><input class="Buttons" Name=<%= sSubmitName %> Value="Upload ESD Files" Type=Submit></td></tr>


</Form>

<SCRIPT>
//Open window with progress bar.
function ProgressBar(){
  var ProgressURL
  ProgressURL = 'progress.asp?UploadID=<%=UploadID%>'

  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200')

  return true;
}
</SCRIPT> 

<% end if

createFoot thisRedirect, 2%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
