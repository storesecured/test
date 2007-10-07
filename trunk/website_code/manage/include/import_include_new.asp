<%

  'Stores only files with size less than MaxFileSize

  Destination_Folder = fn_get_sites_folder(Store_Id,"Upload")
  sOnlyFiles = "safe"
%>
<!--#include file="huge_upload.asp"-->
<%

Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")

  if Form.State <> 0 and Form.State <= 10 then
  	 lsFileName = Trim(Request("DeleteFile"))
  	 If lsFileName <> "" Then
  		on error resume next
  		 Set fso = CreateObject("Scripting.FileSystemObject")
  		 fso.DeleteFile(Upload_Folder&  "\" & lsFileName)
  		 set fso = Nothing
		 fn_redirect "import_"&sImportType&".asp"
  		on error goto 0
  	 End If
	 if request.form("UploadId")<>"" then
		fn_redirect "import_"&sImportType&".asp"
	end if
  end if
sFormAction = "import_"&sImportType&"_action_new.asp"
thisRedirect = "import_"&sImportType&".asp"
sQuestion_Path = "import_export/import_"&sImportType&".htm"
createHead thisRedirect

if Service_Type < 7	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% else %>
</form>

	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="21">
		<%= sThisError %>
			Import <%= sImportType %> from a text file to your store.<br>
		 Using this module you can insert or update <%= sImportType %> in your DB directly from a pre formatted text file.<br><br>
			</td>
	</tr>

	<tr bgcolor='#FFFFFF'>
				<td width="102%" height="38">

						<b>Step 1</b>: Upload <%= sImportType %> file to server:


		 
                        <BR>
                        <a href=include/sample_import/import_<%= sImportType %>_format.asp class=link target=_Blank>File Format Specification</a>
								<BR>
								<a href=other_download.asp?File_Location=sample_location/sample<%= sImportType %>.xls class=link>Sample Excel File</a>
								<BR>
								<a href=other_download.asp?File_Location=sample_location/sample<%= sImportType %>.txt class=link>Sample Tab Delimited File</a>


							<form method="POST" action="<%=PostURL %>" ENCTYPE="multipart/form-data" OnSubmit="return ProgressBar();" >
								<p>Select a file from your local hard drive:
									<input type="file" name="File1"><br>
									<input class="Buttons" type="submit" value="Upload File To Server" name="Upload_Request"></p>
							</form>

							<form method="POST" action="import_<%= sImportType %>_action_new.asp">


								<% 'DISPLAY THE LIST OF CURRENT IMPORTED FILES %>
								<% Dim fso, f, f1, fc %>
								<% Set fso = CreateObject("Scripting.FileSystemObject") %>
								<% Set f = fso.GetFolder(Upload_Folder) %>
								<% Set fc = f.Files	%>
								<% tfiles = 0 %>
							<% For Each f1 in fc %>
							       <% if tfiles=0 then %>
							          <p><b>Step 2:</b> Choose a file from your uploaded files list:</p>

            							<table border="1" cellspacing="0" width=100%>
            						
            								<tr align='center' bgcolor='#0069B2' class='white'>
            							<td><font class=white><b>File name</b></font></td>
            									<td><font class=white><b>Date uploaded</b></font></td>
            									<td><font class=white><b>Select</b></font></td>
            							<td>&nbsp;</td>
            								</tr>
							       <% end if %>
									<tr bgcolor='#FFFFFF'>
								<td><a class="link" href="admin_download.asp?File_Location=<%= Upload_Folder %><%= server.urlencode(f1.name) %>"><%= f1.name %></a></td>
								<td><%= f1.DateLastModified %></td>
								<td><input class="image" type="radio" name="Import_Filename" value="<%= f1.name %>" 
								<% if tfiles = 0 then %>
								checked
								<% end if %>
								></td>
								<td><a class="link" href="import_<%= sImportType %>.asp?DeleteFile=<%= server.urlencode(f1.name) %>">Delete</a></td>
									</tr>
									<% tfiles = tfiles + 1 %>
								<% Next %>

							<% If tfiles>0 then %>
							</table><p>Choose file type:
								<select size="1" name="Delimiter">
							<option value="0">Excel File</option>
						<option value="1">Comma Delimited</option>
						<option value="2">Semicolon Delimited</option>
						<option value="3">Tab Delimited </option>
						<option value="4">Space Delimited</option>
						<option value="5">Other</option>
					</select> Other Delimiter:
								<input type="text" name="Delimiter_Other" size="10"><br><br>
								<input type=hidden name="Qualifier" value="3">
								</p>
					<p><b>Step 3:</b> Press the import button:</p>
								<p align="left">
									<input class="Buttons" type="submit" value="Import <%= sImportType %> File" name="Import_File">
								</p>
							<% End If
							set fso = Nothing
							%>

							</form>

					</td>
				</tr>
			
<% end if %>
<% createFoot thisRedirect, 0%>

<SCRIPT>
//Open window with progress bar.
function ProgressBar(){
  var ProgressURL
  ProgressURL = 'progress.asp?UploadID=<%=UploadID%>'

  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=150')

  return true;
}
</SCRIPT>