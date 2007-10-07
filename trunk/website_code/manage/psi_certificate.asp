<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
on error resume next

Key_Folder = fn_get_sites_folder(store_id,"Key")
if request.form <> "" then
	Set FileObject = CreateObject("Scripting.FileSystemObject")
    File_Full_name = Key_Folder&"\cert.pem"

	 If not FileObject.FolderExists(Key_Folder) Then
		set f1 = FileObject.CreateFolder(Key_Folder)
	 end if

	 FileObject.DeleteFile(File_Full_name)
	 if request.form("Certificate") <> "" then
		Set MyFile = FileObject.OpenTextFile(File_Full_name, 8,true)
		MyFile.Write request.form("Certificate")
		MyFile.Close
		set FileObject = Nothing
	 else
		FileObject.DeleteFolder(Key_Folder)
	 end if
end if

Set FileObject = CreateObject("Scripting.FileSystemObject")
File_Full_name = Key_Folder&"\cert.pem"


Set MyFile = FileObject.OpenTextFile(File_Full_name, 1,true)
Certificate = MyFile.ReadAll
set FileObject = Nothing

sTitle = "PSIGate Instructions"
thisRedirect = "psi_certificate.asp"
sMenu="general"
sFormAction = "psi_certificate.asp"
sName = "PSIGate"
createHead thisRedirect  %>

<TR bgcolor='#FFFFFF'><TD colspan=3>1. Paste a copy of your cert.pem file in the box below<BR>The cert.pem file
		should of been included in your welcome email from PSI Gate, if you cannot find this certificate please contact
		PSI Gate for a replacement.	The certificate file will start with "-----BEGIN RSA PRIVATE KEY-----" and end with "-----END CERTIFICATE-----" and there
		will be random characters in between.	The certificate must be copied into this box exactly, if there are any extra characters or missing
		characters it will not be accepted by PSIGate.<BR>
		2. Select the save button to save changes
		</td></tr>
<TR bgcolor='#FFFFFF'>
	<td width="30%" class="inputname"><b>Certificate File</b></td>
	<td width="70%" class="inputvalue"><textarea rows=20 cols=55 name="Certificate"><%= Certificate %></textarea>
		 <% small_help "Certificate" %></td>
</tr>
<TR bgcolor='#FFFFFF'><TD colspan=3><a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
		</td></tr>

<% createFoot thisRedirect, 1 %>
