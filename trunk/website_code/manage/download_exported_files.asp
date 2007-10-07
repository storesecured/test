<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
lsFileName = Trim(Request("File"))
Export_Folder = fn_get_sites_folder(Store_Id,"Export")
If lsFileName <> "" Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(Export_Folder & lsFileName)
End If

sFormAction = "download_exported_files.asp"
sTitle = "Download Exported Files"
sFullTitle = "My Account > Download Exported Files"
thisRedirect = "download_exported_files.asp"
sMenu="account"
sQuestion_Path = "import_export/export.htm"
createHead thisRedirect
if Service_Type < 7 then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% else %>
<SCRIPT LANGUAGE="Javascript">

	<!--
	function VerifyDelete(ID)
	{
		return confirm("Are you sure you want to delete this file?");}
	// -->
</script>


	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="28">
			<input type="button" class="Buttons" value="Export Orders" name="Back_To_Item" OnClick=JavaScript:self.location="orders.asp" >
			<input type="button" class="Buttons" value="Export Customers" name="Back_To_Item" OnClick=JavaScript:self.location="My_Customer_Base.asp" >
			<input type="button" class="Buttons" value="Export Items" name="Back_To_Item" OnClick=JavaScript:self.location="edit_items.asp" >
			<input type="button" class="Buttons" value="Export Shipping" name="Back_To_Item" OnClick=JavaScript:self.location="export_shipping.asp" >
			<input type="button" class="Buttons" value="Export Departments" name="Back_To_Item" OnClick=JavaScript:self.location="export_departments.asp" >
			<input type="button" class="Buttons" value="Export Newsletter Subscribers" name="Back_To_Item" OnClick=JavaScript:self.location="export_newsletter.asp" >
			</td>
	</tr>

	<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='13'>
		
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<tr bgcolor='#FFFFFF'><td colspan=4>Click on a filename to download the file.</td></tr>
			 <tr align='center' bgcolor='#0069B2' class='white'>
					<td colspan="2"><font class=white><b>Filename</b></font></td>
					<td><font class=white><b>Created</b></font></td>
			 <td><font class=white><b>Delete</b></font></td>
			</tr>
				<% Dim fso, f, f1, fc %>
				<% Set fso = CreateObject("Scripting.FileSystemObject") %>
				<% Set f = fso.GetFolder(Export_Folder) %>
				<% Set fc = f.Files %>
			<% For Each f1 in fc %>
					<% if not (Ucase(f1.name) = Ucase("CCMck.log") or	Ucase(f1.name) = Ucase("merchant_conf.inc")) then %>
						<tr bgcolor='#FFFFFF'>
					<td colspan="2">
								<a class="link" href="admin_download.asp?File_Location=<%= Export_Folder %>\<%= f1.name %>"><%= f1.name %></a></td><td><%= f1.DateLastModified %></td>
							<td>
								<a class="link" href="download_exported_files.asp?File=<%= Server.URLEncode(f1.name) %>" OnClick="return VerifyDelete('0');">Delete</a></td>
				</tr>
					<% end if %>
				<% Next %>	
		</table>
		</td>
	</tr>
<% end if %>


<% createFoot thisRedirect, 0%>

<%
if trim(request.querystring("new")) <> "" then
efile = trim(request.querystring("new"))
%>
<SCRIPT LANGUAGE="Javascript">
<!--
alert('Data exported to file <%=efile%>.txt')
// -->
</script>
<%
end if
%>