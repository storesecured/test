<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%
on error goto 0
if request.form <> "" then
	Name = checkStringForQ(request.form("Name"))
	Detail = replace(request.form("Detail"),"'","''")
	Support_Id = checkStringForQ(request.form("Support_Id"))
	Email = checkStringForQ(request.form("Email"))
	Status = checkStringForQ(request.form("Status"))
	IP_Address= Request.ServerVariables("REMOTE_ADDR")
	
	Support_Path=request.form("Support_Path")


	sql_select = "select assigned_to from Sys_Support where support_Id="&Support_Id
	set rs_store=conn_store.Execute(sql_select)
	nAdmin=rs_store("assigned_to")

	select case nAdmin
		case 1: Admin_Email="melanie@easystorecreator.com"
		case 2: Admin_Email="abbas@easystorecreator.com"
		case 3: Admin_Email="chris@easystorecreator.com"
		case 4: Admin_Email="ganesh@easystorecreator.com"
		case 5: Admin_Email="cheriyan@easystorecreator.com"
		case 6: Admin_Email="umesh@easystorecreator.com"
		case else: Admin_Email=sNoReply_email
	end select

	sql_insert = "Add_Support_Request1 "&Store_Id&",'"&Detail&"','"&Name&"','"&Email&"',"&Support_Id&",DEFAULT,DEFAULT,'"&IP_Address&"','"&Support_Path&"'"

	conn_store.Execute sql_insert
	if Support_Path="" then
		AttachmentFile=NULL
	else
		AttachmentFile= fn_get_sites_folder(Store_Id,"Root") &"\"&Support_Path
	end if
	Send_Mailsupport  sNoReply_email,Admin_Email,"Support Request "&Store_Id&" "&Subject& "Updated",Detail&vbcrlf&"http://manage.easystorecreator.com/support_edit_admin.asp?Support_Id="&Support_Id , AttachmentFile
	response.redirect "support_list.asp"
end if

sTitle = "Edit Support Request"
sFullTitle = "My Account > <a href=support_list.asp class=white>Support Request</a> > Edit"
sCancel="support_list.asp"
sCommonName="Support Request"
sFormAction = "support_edit.asp"
thisRedirect = "support_request.asp"
sMenu = "account"
sQuestion_Path = "import_export/my_account/support_requests.htm"
createHead thisRedirect

Support_Id = request.querystring("Id")

if Support_Id <> "" and isNumeric(Support_Id) then
	sql_select = "select support_path,Subject, Status,sys_support_detail.sys_created,detail,sys_support_detail.poststatus,created_by,store_id,email from Sys_Support inner join sys_support_detail on sys_support.support_id = sys_support_detail.support_id where sys_support.support_Id="&Support_Id&" and store_id="&Store_Id&" order by sys_support_detail.sys_created"

	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	if request.querystring("ViewOnly") <> "" then
		fViewOnly = 1
	else
		fViewOnly = 0
	end if
else
	response.redirect "support_list.asp"
end if
%>


				<tr bgcolor='#FFFFFF'>
					 <td colspan="3"><input class=buttons type="button" value="Close Support Request" OnClick=JavaScript:self.location="support_list.asp?Support_Id=<%=Support_Id %>&Status=4"></td>
				 </tr>
				 <%
				 str_class=1
				 if noRecords = 0 then
				FOR rowcounter= 0 TO myfields("rowcount")
					if str_class=1 then
					  str_class=0
						 else
						str_class=1
						 end if
					iStatus = mydata(myfields("status"),rowcounter)
					if iStatus = 1 then
						sStatus = "New"
					elseif iStatus = 2 then
						sStatus = "Assigned"
					elseif iStatus = 3 then
						sStatus = "Cancelled"
					elseif iStatus = 4 then
						sStatus = "Completed"
					elseif iStatus = 5 then
						sStatus = "Reopened"
					end if 
	
				
					vStatus = mydata(myfields("poststatus"),rowcounter)
					if vStatus = 1 then
						pStatus = "New"
					elseif vStatus = 2 then
						pStatus = "Assigned"
					elseif vStatus = 3 then
						pStatus = "Cancelled"
					elseif vStatus = 4 then
						pStatus = "Completed"
					elseif vStatus = 5 then
						pStatus = "Reopened"
					elseif vStatus = 6 then
						pStatus = "Awaiting Feedback"
					end if
					support_path= mydata(myfields("support_path"),rowcounter)
					%>
				<tr bgcolor='#FFFFFF'><TD colspan=2 class="<%= str_class %>" size='100%'><B>At&nbsp;<%= mydata(myfields("sys_created"),rowcounter) %>&nbsp;<%= mydata(myfields("created_by"),rowcounter) %>&nbsp;wrote:</b></td><td class="inputvalue" align="center" size=5><B><a class="help" onmouseover="return overlib('Status When Posted:&nbsp;<%=pStatus%>');" onmouseout="return nd();">(I)</a>&nbsp;&nbsp;</b>
				</td></tr>
				<tr bgcolor='#FFFFFF'><TD colspan=3 class="<%= str_class %>"><%= replace(mydata(myfields("detail"),rowcounter),vbcrlf,"<BR>") %>
            <% if Support_Path <> "" then %>
            <BR><B>Attachment:</b> <a href=<%= Site_Name %><%=support_path%> target=_blank class=link><%=support_path%></a>

            <% end if %>
            </td></tr>
				

				<% 
            if rowcounter=0 then
            name = mydata(myfields("created_by"),rowcounter)
				email = mydata(myfields("email"),rowcounter)
				status = mydata(myfields("status"),rowcounter)
				end if
				next
				else
				    response.redirect "admin_home.asp"
				end if
				if fViewOnly <> 1 then %>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Name</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Name" value="<%= name %>" size="60">
					<INPUT type="hidden"  name="Name_C" value="Re|String|0|150|||Name">
					<% small_help "Name" %></td>
				</tr>
				
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Detail</b></td>
					<td width="90%" class="inputvalue"><textarea name="Detail" rows=10 cols=55></textarea>
					<INPUT type="hidden"  name="Detail_C" value="Re|String|0|3000|||Detail">
					<% small_help "Detail" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname">Attachment</td>
					<td width="70%" class="inputvalue">
						<input type="text" name="Support_Path" value="" size="60">
						<input type="hidden" name="Support_Path_C" value="Op|String|0|50|||Support Path">
						<font size="1" color="#000080">
						<a class="link" href="JavaScript:goSupportFileUploader('Support_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="File Upload"></a>
						<% small_help "Attachment" %>
					</td>
				</tr>

				<input type="hidden" name="Support_Id" value="<%= Support_Id %>">
				<input type="hidden" name="Status" value="<%= status %>">
				<input type="hidden" name="Email" value="<%= email %>">
				<% createFoot thisRedirect, 1%>
				<% else %>
<% createFoot thisRedirect, 0%>
<% end if %>
