<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
calendar=1
'CODE FOR DELETING A LABEL
Label_Del = Request.QueryString("Label_Del")
if Label_Del <> "" then
	'DELETE THE FILE FROM DISK
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")
	File_Full_Name = Upload_Folder&"/ups_accept_op_"& Label_del&".xml"
	On Error Resume Next
	FileObject.DeleteFile(File_Full_Name)
	Set FileObject = Nothing
end if


if  Request.Form("Delete_Item") <> "" then
	if Request.Form("Delete_Item")="SEL" then
	  if request.form("DELETE_IDS") <> "" then
		  Set FileObject = CreateObject("Scripting.FileSystemObject")
		  for each this_file_name in Split(request.form("DELETE_IDS"), ",")
		  this_file_name = replace(this_file_name," ","")
			if this_file_name <> "" then
				File_Full_Name=Upload_Folder&this_file_name
				On Error resume next
				FileObject.DeleteFile(File_Full_Name)
			end if
		  next
		  Set FileObject = Nothing
		end if
	end if
end if

on error resume next

sFormName ="Print_Labels"
sFormAction = "labels_list.asp"
sTitle = "Label List"
thisRedirect = "labels_list.asp"
sMenu = "Advanced"
createHead thisRedirect
	

%>


<SCRIPT LANGUAGE = "JavaScript">

	  function goDeleteSeleted(){
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked)
				tsel = tsel+1;

		if (confirm("Are you sure you want to delete these Labels?")){
			document.forms[0].Delete_Item.value="SEL";
			document.forms[0].submit(); }}

		function selectAllItems(){
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
				document.forms[0].DELETE_IDS[i].checked = document.forms[0].selectall.checked; }



		//function for deleting labels
		function ConfirmDelete(stype, objName)
		{
			 if(stype.toUpperCase() == "FILE")
			{
				if (confirm("Are you sure you want to delete this Label?"))
					document.location.href="labels_list.asp?Label_Del="+objName+"&<%=sAddString%>"
			}
		}

</SCRIPT>
<input type="Hidden" name="Delete_Item" value="">

	<tr>
	<td width="200" height="74" align="center" valign=top>
     <%
			strHTML = strHTML & ("<table width='100%' class=list>")
			
				Dim fso, f, f1, fc
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.GetFolder(Upload_Folder)
				Set fc = f.Files
				Set ff = f.SubFolders
				tfiles = 0
				str_class=1
			
				For Each f1 in fc

  					  if str_class=1 then
  					  str_class=0
  						 else
  						str_class=1
  						 end if
	  	if instr(f1.type,"XML") > 0 and instr(f1.name, "ups_accept_op_") > 0 then

						ext=	split(f1.name, "_")
						fileext=ext(Ubound(ext))
						orderid=split(fileext, ".")

					   strHTML = strHTML & ("<tr class='"&str_class&"'>")
  
  						if instr(f1.type,"XML") > 0 and instr(f1.name, "ups_accept_op_") > 0 then
							strHTML = strHTML & ("<td> Order # "&orderid(0)&"</td>")
							strHTML = strHTML & ("<td> "&f1.datelastmodified&"</td>")
							strHTML = strHTML & ("<td><a class='link' target='_new' href=""ups_print_labels_out.asp?id="& orderid(0) & """>Print</a>")
							strHTML = strHTML & ("</td><td><a class='link' target='_new' href=""ups_cancel_labels.asp?id="& orderid(0) & """>Cancel</a>")
  						else
  							strHTML = strHTML & ("<td><img src=images/file.gif></td><td>"&f1.name&"</td><td>")
  
					 end if
  						strHTML = strHTML & ("</td><td align='center'><a class='link' href=""JavaScript:ConfirmDelete('FILE','"&orderid(0)&"');"">Delete</a> </td>")
  						strHTML = strHTML & ("<td align=center><input class='image' type='checkbox' name='DELETE_IDS' value="&f1.name&">")
			end if
  				tfiles = tfiles + 1
				Next

				if tfiles > 1 then
					strHTML = strHTML & ("<tr><td colspan=4></td><td align='center'>")
					strHTML = strHTML & ("<input class='image' type='checkbox' name='selectall' onClick=""JavaScript:selectAllItems();""><br>Check All</td></tr>")
					strHTML = strHTML & ("<tr><td colspan=4 align=center><input type='Button' name='DeleteAll' value='Delete Selected' onClick=""JavaScript:goDeleteSeleted();""></td></tr>")
				end if
			
		strHTML = strHTML & ("</table>") %>

		<table border=0 cellpadding=0 cellspacing=0 width=190 ID="Table4">
				<tr><td>
				<div class="bar" style="padding-left: 5px; padding-top: 10px">
				<font size="2" face="tahoma" color="black"><b>Order ID</b></font>
				</div>
				</td>
				
				<td>
				<div class="bar" style="padding-left: 0px; padding-top: 10px">
				<font size="2" face="tahoma" color="black"><b>Created</b></font>
				</div>
				</td>
				</tr>
				</table>
				
				<div style="overflow:auto;height:320;width:360;BORDER-LEFT: #316AC5 1px solid;BORDER-RIGHT: LightSteelblue 1px solid;BORDER-BOTTOM: LightSteelblue 1px solid;">
				<%=strHTML%>
				</div>
		  </table>
  	</td>
	</tr>


 


