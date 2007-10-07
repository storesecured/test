<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%
on error goto 0
if request.form <> "" then
	Subject = checkStringForQ(request.form("Subject"))
	Name = checkStringForQ(request.form("Name"))
	Email = checkStringForQ(request.form("Email"))
	Detail = nullifyQ(request.form("Detail"))
	IP_Address= Request.ServerVariables("REMOTE_ADDR")
	Support_Path=request.form("Support_Path")

dim strstat
	sql_insert = "New_Support_Request1 "&Store_Id&",'"&Subject&"','"&Detail&"','"&Name&"','"&Email&"','"&IP_Address&"','"&Support_Path&"'"
	conn_store.Execute sql_insert

	if Support_Path="" then
		AttachmentFile=NULL
	else
		AttachmentFile= fn_get_sites_folder(Store_Id,"Root") & Support_Path
	end if
	response.redirect "support_list.asp"
end if

sInstructions="Before requesting support you might want to check the <a href=http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/front_page!PAGETYPE class=link target=_blank>Help Section</a> to see if you question is already answered.  Note that when using this support request form your store id will be automatically captured.  If you choose to contact support in another manner please make sure to reference your store id: "&Store_Id
             
sTitle = "Add Support Request"
sFullTitle = "My Account > <a href=support_list.asp class=white>Support Request</a> > Edit"
sCancel="support_list.asp"
sCommonName="Support Request"
sFormAction = "support_request.asp"
thisRedirect = "support_request.asp"
sMenu = "account"
sQuestion_Path = "import_export/my_account/support_requests.htm"
createHead thisRedirect
%>


				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Name</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Name" value="" size="60" maxlength=150>
					<INPUT type="hidden"  name="Name_C" value="Re|String|0|150|||Name">
					<% small_help "Name" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Email</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Email" value="<%= Store_Email %>" size="60" maxlength=150>
					<INPUT type="hidden"  name="Email_C" value="Re|Email|0|150|||Email">
					<% small_help "Email" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Subject</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Subject" value="" size="60" maxlength=150>
					<INPUT type="hidden"  name="Subject_C" value="Re|String|0|150|||Subject">
					<% small_help "Subject" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Detail</b></td>
					<td width="90%" class="inputvalue">
          <input readonly type=text name=remLenDet size=3 class=char maxlength=3 value="<%= 3000-len(Detail) %>" class=image><font size=1><I>characters left</i></font>
          <BR><textarea name="Detail" rows=10 cols=55 onKeyDown="textCounter(this.form.Detail,this.form.remLenDet,3000);" onKeyUp="textCounter(this.form.Detail,this.form.remLenDet,3000);"></textarea>
					<INPUT type="hidden"  name="Detail_C" value="Re|String|0|3000|||Detail">
					<% small_help "Detail" %></td>
				</tr>

<tr bgcolor='#FFFFFF'>
	<td width="30%" class="inputname">Attachment</td>
	<td width="70%" class="inputvalue">
		<input type="text" name="Support_Path" value="<%=Support_Path%>" size="60">
		<input type="hidden" name="Support_Path_C" value="Op|String|0|50|||Support Path">
		<font size="1" color="#000080">
		<a class="link" href="JavaScript:goSupportFileUploader('Support_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="File Upload"></a>
		<% small_help "Attachment" %>
	</td>
</tr>


<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Name","req","Please enter your name.");
 frmvalidator.addValidation("Email","req","Please enter a valid email.");
 frmvalidator.addValidation("Subject","req","Please enter a subject.");
 frmvalidator.addValidation("Detail","req","Please enter the details of your support request.");

</script>
