<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

'CODE FOR DELETING AN IMAGE
File_Del = Request.QueryString("File_Del")
Download_Folder = fn_get_sites_folder(Store_Id,"Download")

if File_Del <> "" then
	'DELETE THE FILE FROM DISK
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	File_Name = Download_Folder&"\"&File_Del
	on error resume next
	FileObject.DeleteFile(File_Name)
end if

sInstructions="ESD stands for electronic software download.  ESD files are files which are to be delivered to your customer at checkout.  This can be things like ebooks, software programs, work documents, images, zip files, pdf files, or anything electronic that needs to be delivered on successfull purchase."

sTitle = "ESD Files"
sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > ESD Files"
thisRedirect = "file_picker.asp"
sMenu = "inventory"
sQuestion_Path = "inventory/file_list.htm"
createHead thisRedirect
if Service_Type < 7 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect,0 %>

<%
elseif Trial_Version then %>
  <TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available for trial stores.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% else %>



	<TR bgcolor='#FFFFFF'>
		<form action="" method="post">	
		<td colspan="3" height="23">
			<input type="button" class="Buttons" value="Upload ESD Files" name="Add" OnClick=JavaScript:self.location="upload_files.asp?<%= sAddString %>" ><br>

		</td>
		</form>
	 
	</tr>

	<TR bgcolor='#FFFFFF'>
	

					<td><b>Name</b></td>
					<td align="center"><b>Delete</b></td>

				</tr>
		
				<%
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
					<tr bgcolor="#FFFFFF">
						<td><%= f1.name %></td>


						<td align="center"><a class="link" href="file_list.asp?File_Del=<%= f1.name %>&<%= sAddString %>">Delete</a> </td></tr>

				<% tfiles = tfiles + 1
				Next %>

	


<% end if %>
<% createFoot thisRedirect, 0%>

