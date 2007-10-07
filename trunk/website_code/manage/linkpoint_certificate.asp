<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
on error resume next

if request.form <> "" then
	Set FileObject = CreateObject("Scripting.FileSystemObject")
    Key_Folder = fn_get_sites_folder(Store_Id,"Key")
    Key_File = Key_Folder&"\cert.pem"

	 If not FileObject.FolderExists(Key_Folder) Then
		set f1 = FileObject.CreateFolder(Key_Folder)
	 end if

	 FileObject.DeleteFile(Key_File)
	 if request.form("Certificate") <> "" then
		Set MyFile = FileObject.OpenTextFile(Key_File, 8,true)
		MyFile.Write request.form("Certificate")
		MyFile.Close
		set FileObject = Nothing
	 else
		FileObject.DeleteFolder(Key_Folder)
	 end if
end if

Set FileObject = CreateObject("Scripting.FileSystemObject")

Set MyFile = FileObject.OpenTextFile(Key_File, 1,true)
Certificate = MyFile.ReadAll
set FileObject = Nothing

sTitle = "Linkpoint Instructions"
thisRedirect = "linkpoint_certificate.asp"
sMenu="general"
sFormAction = "linkpoint_certificate.asp"
sName = "Linkpoint"
createHead thisRedirect  %>

<TR bgcolor='FFFFFF'><TD colspan=3>1. When signing up select either LinkPoint API or LinkPoint Central<BR>
		2. Paste a copy of your cert.pem file in the box below<BR>The cert.pem file
		should of been included in your welcome email from linkpoint, if you cannot find this certificate please contact
		linkpoint for a replacement.	The certificate file will start with "-----BEGIN RSA PRIVATE KEY-----" and end with "-----END CERTIFICATE-----" and there
		will be random characters in between.	The certificate must be copied into this box exactly, if there are any extra characters or missing
		characters it will not be accepted by Linkpoint.  If formatted properly your certificate file 
      should fit in the box below with no word wrap or scrolling needed.<BR>
		3. Select the save button to save changes
		</td></tr>
<TR bgcolor='#FFFFFF'>
	<td width="100%" class="inputname"><b>Certificate File</b><BR>
	<textarea rows=32 cols=65 name="Certificate"><%= Certificate %></textarea>
		 <% small_help "Certificate" %></td>
</tr>
<TR bgcolor='#FFFFFF'><TD colspan=3><a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
		</td></tr>

<% createFoot thisRedirect, 1 %>
