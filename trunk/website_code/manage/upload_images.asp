<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
sOpenPath = request.querystring("Open")
sOpenPath = replace(replace(request.querystring("Open"),"#","\"),"..\","")

Destination_Folder = fn_get_sites_folder(Store_Id,"Images")&sOpenPath&"\"
sOnlyFiles = "safe"
%>
<!--#include file="include/huge_upload.asp"-->
<%

  if sOpenPath <> "" then
     PostURL = PostURL & "&Open="&server.urlencode(sOpenPath)
  end if

sFlashHelp="irfanview/irfanview.htm"
sMediaHelp="irfanview/irfanview.wmv"
sZipHelp="irfanview/irfanview.zip"
sInstructions="Do not upload images which contain special characters such as the ampersand, comma, space, colon, semicolon, brackets, parenthesis, etc.<BR><BR>You may also upload and link to pdf, htm, css or other safe file types for display on your website.  All potentially dangerous file types are blocked from upload."
sTextHelp="images/upload.doc"


if request.querystring("UploadID")<>"" then
        response.redirect "left_image_picker.asp"
end if

sName = "upload_image_form"
sTitle = "Upload Files/Images"
sFullTitle = "Design > <a href=left_image_picker.asp class=white>Files/Images</a> > Upload"
sSubmitName = "Upload_Image"
thisRedirect = "upload_images.asp"
sEncType = "ENCTYPE='multipart/form-data'"
sMenu="design"
sQuestion_Path = "images/upload_images.htm"
createHead thisRedirect

if SizeUsage > 100 then %>
	<TR bgcolor='#FFFFFF'><td><font color=red>You are not allowed to upload anything because you are currently over your size limit.  You are using (<%=SizeUsage %>% of your available space.).
	<BR><BR>
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</font></td></tr>
<%
else

%>
</form><form method="POST" action="<%=PostURL %>" ENCTYPE="multipart/form-data" OnSubmit="return ProgressBar();" >

				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image1" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image2" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image3" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image4" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image5" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image6" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image7" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image8" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image9" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="8%"><B>Filename</B></td><td class="inputvalue" width="92%"><input type="file" name="Image10" size="60"><% small_help "Filename" %></td></tr>
				<TR bgcolor='#FFFFFF'><TD width="100%" colspan=3 align=center><input type="submit" class="Buttons" value="Upload Files/Images" name="Upload_Image"></td>
				</tr>
<%
end if

createFoot thisRedirect, 2%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
<SCRIPT>
//Open window with progress bar.
function ProgressBar(){
  var ProgressURL
  ProgressURL = 'progress.asp?UploadID=<%=UploadID%>'

  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=150')

  return true;
}
</SCRIPT>
