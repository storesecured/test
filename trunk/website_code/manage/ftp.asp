<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sFlashHelp="smartftp/ftp.html"
'sMediaHelp="smartftp/smartftp.wmv"
'sZipHelp="smartftp/smartftp.zip"

sTitle = "FTP"
sFullTitle = "General > FTP"
thisRedirect = "ftp.asp"
sMenu="general"
addPicker=1
sQuestion_Path = "general/ftp.htm"
createHead thisRedirect
if Trial_Version then %>
	<TR bgcolor=FFFFFF>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% else 
sql_select = "SELECT Store_User_Id,Store_Password FROM Store_Settings WHERE Store_id="&Store_id
rs_store.open sql_select,conn_store,1,1
Username=rs_store("Store_User_Id")
Password=rs_Store("Store_Password")
rs_store.close
%>


		<TR bgcolor='#FFFFFF'><TD>

		 <B>FTP Configuration</b><BR>
		 FTP Server = <%= ftp_server %><BR>
		 Account Name = <%= Username %><BR>
		 Password = <%= Password %><BR>
		 Port = <%= ftp_port %>


		 <BR><BR><B>FTP Program</B><BR>
		 To access your ftp account you may use any FTP program.  If you do not already have one we recommend using
		<a class="link" href="http://www.smartftp.com/?affiliate=blac6789" target="_blank">SmartFTP</a><BR>


	  <BR><B>FTP Directories</b><BR>
	  <I>Website Files</i> = This is the directory where any root files should go such as html files<BR>
	  <I>Website Files/Images</i> = This is the directory where anything the should be displayed on your website will go, ie images, css, etc.<BR>
	  <I>Hidden Admin Files/ESD_Download</I> = This is the directory where you would upload any software, music, etc that you are selling as downloadable content.<BR>
     <I>Hidden Admin Files/Upload</i> = This is the directory where you would upload files that are to be used for updating inventory, coupons, customers, etc.<BR>
     <I>Hidden Admin Files/Export</i> = This is the directory where you can download exported files including orders, inventory, etc.

     </td></tr>

	  <TR bgcolor=FFFFFF><TD></td></tr>
<% end if %>
<% createFoot thisRedirect, 0 %>


