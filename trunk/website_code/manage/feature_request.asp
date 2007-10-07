<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
if request.form <> "" then
	Subject = checkStringForQ(request.form("Subject"))
	Detail = checkStringForQ(request.form("Detail"))
	Area = checkStringForQ(request.form("Area"))
	Section = checkStringForQ(request.form("Section"))
	sType = checkStringForQ(request.form("Time"))
	sql_insert = "insert into sys_features (store_id, subject, detail, status, area, area2, feature_type) values ("&Store_Id&",'"&Subject&"','"&Detail&"','New','"&Area&"','"&Section&"','"&sType&"')"
   conn_store.Execute sql_insert
   if Store_Id > 120 then
   	'store less then 120 are admin stores
   	Send_Mail Store_Email,sDeveloper_email,"Feature Request "&Store_Id&" "&Subject&"-"&sType,Detail
   end if
	response.redirect "features2.asp"
end if
sInstructions="Requests for new features will be recorded for future implementation and you will be notified if one of your requested features is implemented.  There is no guarantee that your request will be implemented but it will be considered.  Please do not use this area to ask questions."

addPicker = 1
sTitle = "Add Feature Request"
sFullTitle = "My Account > <a href=features2.asp class=white>Feature Requests</a> > Add"
sCommonName="Feature"
sCancel="features2.asp"
sFormAction = "feature_request.asp"
thisRedirect = "feature_request.asp"
sMenu = "account"
createHead thisRedirect
%>


				
			    	<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Subject</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Subject" value="" size="60" maxlength=150>
					<INPUT type="hidden"  name="Subject_C" value="Re|String|0|150|||Name">
					<% small_help "Subject" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Detail</b></td>
					<td width="90%" class="inputvalue">
          <input readonly type=text name=remLenDet size=3 class=char maxlength=3 value="<%= 3000-len(Detail) %>" class=image><font size=1><I>characters left</i></font>
          <BR><textarea name="Detail" rows=5 cols=55 onKeyDown="textCounter(this.form.Detail,this.form.remLenDet,3000);" onKeyUp="textCounter(this.form.Detail,this.form.remLenDet,3000);"></textarea>
					<INPUT type="hidden"  name="Detail_C" value="Re|String|0|3000|||Detail">
					<% small_help "Detail" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Area</b></td>
					<td width="90%" class="inputvalue"><select name=Area><option value=Admin>Admin
					<option value=Store>Store</select>
					<% small_help "Area" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Section</b></td>
					<td width="90%" class="inputvalue"><select name=Section><option value=Design>Design
					<option value=Orders>Orders<option value=Settings>Settings<Option value=Inventory>Inventory<option value=Marketing>Marketing<option value=Reports>Reports<option value=Import/Export>Import/Export<option value=Emails>Emails<option value=Other>Other</select>
					<% small_help "Section" %></td>
				</tr>
				<input type=hidden name=Time value="Free">

				
				

				

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Subject","req","Please enter a subject.");
 frmvalidator.addValidation("Detail","req","Please enter the details of this feature request.");


</script>
