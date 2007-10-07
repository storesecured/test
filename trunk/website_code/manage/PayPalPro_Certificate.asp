<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
on error resume next

Key_Folder = fn_get_sites_folder(store_id,"Key")
if request.form <> "" then
	Set FileObject = CreateObject("Scripting.FileSystemObject")
    File_Full_Name= Key_Folder&"cert_key_pem.txt"
	 If not FileObject.FolderExists(Key_Folder) Then
		set f1 = FileObject.CreateFolder(Key_Folder)
	 end if

	 FileObject.DeleteFile(File_Full_Name)
	 if request.form("Certificate") <> "" then
		Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
		MyFile.Write request.form("Certificate")
		MyFile.Close
		set FileObject = Nothing
	 else
		FileObject.DeleteFolder(Key_Folder)
	 end if
end if

Set FileObject = CreateObject("Scripting.FileSystemObject")
File_Full_Name = Key_Folder&"cert_key_pem.txt"

Set MyFile = FileObject.OpenTextFile(File_Full_Name, 1,true)
Certificate = MyFile.ReadAll
set FileObject = Nothing

sTitle = "PayPalPro Instructions"
thisRedirect = "PayPalPro_Certificate.asp"
sMenu="general"
sFormAction = "PayPalPro_Certificate.asp"
sName = "PayPal-Pro"
createHead thisRedirect  %>

<TR bgcolor='#FFFFFF'><TD colspan=3>1. Paste a copy of your cert_key_pem.txt file in the box below<BR>The certificate file
		should have been downloaded from PayPal, if you cannot find this certificate please contact
		PayPal .	The certificate file will start with "-----BEGIN RSA PRIVATE KEY-----" and end with "-----END CERTIFICATE-----" and there
		will be random characters in between.	The certificate must be copied into this box exactly, if there are any extra characters or missing
		characters it will not be accepted by PayPal.<BR>
		2. Select the save button to save changes
		</td></tr>
<TR bgcolor='#FFFFFF'>
	<td width="30%" class="inputname" colspan=2><b>Certificate File</b></td></tr>
	<TR bgcolor='#FFFFFF'><td width="70%" class="inputvalue"><textarea rows=20 cols=75 name="Certificate"><%= Certificate %></textarea>
		 <% small_help "Certificate" %></td>
</tr>
<TR bgcolor='#FFFFFF'><TD colspan=3><a href=real_time_settings.asp<%= sAddString %> class=link>Click here to return to payment processor page</a>
		</td></tr>

<% createFoot thisRedirect, 1 %>
