<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<script language="javascript">
function showHideRes()
{
	if (document.getElementById("Status").value==2)
		document.getElementById("divassign").style.display="block";
	else
		document.getElementById("divassign").style.display="none";
}
</script>

<%
on error goto 0
Support_Path=""
if request.form <> "" then
	Store_Id=checkStringForQ(request.form("Store_Id"))
	Name = checkStringForQ(request.form("Name"))
	Detail = replace(request.form("Detail"),"'","''")
	Email = checkStringForQ(request.form("Email"))
	Support_Id = checkStringForQ(request.form("Support_Id"))
	Status = checkStringForQ(request.form("Status"))
	Assigned_to =checkStringForQ(request.form("Assigned_to"))
	IP_Address= Request.ServerVariables("REMOTE_ADDR")

Support_Path= trim(request.form("Support_Path"))
	if Assigned_to <> ""  and Status=2 then
		sql_insert = "Add_Support_Request1 "&Store_Id&",'"&Detail&"','"&Name&"','"&Email&"',"&Support_Id&","&Status&","&Assigned_to&", '"&IP_Address&"','"&Support_Path&"','"&Status&"'"
	else
		sql_insert = "Add_Support_Request1 "&Store_Id&",'"&Detail&"','"&Name&"','"&Email&"',"&Support_Id&","&Status&",DEFAULT,'"&IP_Address&"','"&Support_Path&"','"&Status&"'"
	end if

	conn_store.Execute sql_insert

	strbody=""

		sql_select = "select server,support_path,Subject, Status,sys_support_detail.sys_created,detail,sys_support_detail.IP_Address,sys_support_detail.poststatus,created_by,store_settings.store_id,email from Sys_Support inner join sys_support_detail on sys_support.support_id = sys_support_detail.support_id inner join store_settings on sys_support.store_id=store_settings.store_id where sys_support.support_Id="&Support_Id&" order by sys_support_detail.sys_created"
		set newfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_select,newdata,newfields,noRecords)


		 if noRecords = 0 then
			FOR rowcounter= 0 TO newfields("rowcount")
			    create_date =newdata(newfields("sys_created"),rowcounter) 
			    creator = newdata(newfields("created_by"),rowcounter) 
			    description =newdata(newfields("detail"),rowcounter)
                sites_server =newdata(newfields("server"),rowcounter)
                sites_store_id =newdata(newfields("store_id"),rowcounter)

			    strbody= strbody & vbcrlf & vbcrlf & "At " & create_date & " " & creator & " wrote: " & vbcrlf & description &vbcrlf & "*****************************************************" & vbcrlf
			next
		end if
		
		sSupportUrl= "http://manage.storesecured.com/"

	if request.form("send_email") = 1 and Status=2 then
		admin_mail=""

      admin_mail = ", "&assigned_to&"@storesecured.com"
	

		Email= Email & admin_mail
	end if
	if request.form("send_chatters")=1 then
	   Email =Email & ", joe@storesecured.com, jason@storesecured.com, sam@storesecured.com, cheryl@storesecured.com"
	end if
	
	if Support_Path="" then
		AttachmentFile=NULL
	else
		AttachmentFile= fn_get_sites_folder(Store_Id,"Root") &"\"&Support_Path
	end if
	
	Send_Mail sNoReply_email,Email,"Support Request Updated Store Id:"&sites_store_id,strbody&vbcrlf&"*****************************************************"&vbcrlf&"To respond to this support request go to http://manage.storesecured.com/support_edit.asp?op=edit&id="&Support_Id& vbcrlf&"Do not reply to this email, this is an outgoing mailbox only for automatic notifications."
	response.redirect "support_list_admin.asp"
end if

sTitle = "Support Request Edit - "&Store_Id
sCancel = "support_list_admin.asp"
sFullTitle = "Admin > <a href=support_list_admin.asp class=white>Support Requests</a> > Edit"
sCommonName = "Support Request"
sFormAction = "support_edit_admin.asp"
thisRedirect = "support_request.asp"
sMenu = "account"
createHead thisRedirect

if request.querystring("Support_Id") <> "" then
	Support_Id = request.querystring("Support_Id")
	sql_select = "select support_path,Subject, Status,sys_support_detail.sys_created,site_name,detail,sys_support_detail.IP_Address,sys_support_detail.poststatus,created_by,store_settings.store_id,email from Sys_Support inner join sys_support_detail on sys_support.support_id = sys_support_detail.support_id  inner join store_settings on sys_support.store_id=store_settings.store_id where sys_support.support_Id="&Support_Id&" order by sys_support_detail.sys_created"
	set myfields=server.createobject("scripting.dictionary")
   Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	if request.querystring("ViewOnly") <> "" then
		fViewOnly = 1
	else
		fViewOnly = 0
	end if

else
	response.redirect "support_list_admin.asp"
end if
%>

				<tr bgcolor='#FFFFFF'>
					 <td colspan="4"><input type="button" value="Back to Support List" OnClick=JavaScript:self.location="support_list_admin.asp"></td>
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
					

					

					ip_addr=mydata(myfields("ip_address"),rowcounter)
					support_path= mydata(myfields("support_path"),rowcounter)
					clinetStore_id=mydata(myfields("store_id"),rowcounter)
					site_name = mydata(myfields("site_name"),rowcounter)
				   'Response.Write "clinetStore_id=" & clinetStore_id
					'Response.End 
					%>


				<tr bgcolor='#FFFFFF'><TD colspan=3 class="<%= str_class %>"><B>At&nbsp;<%= mydata(myfields("sys_created"),rowcounter) %>&nbsp;<%= mydata(myfields("created_by"),rowcounter) %>&nbsp;wrote:</b></td><td class="inputvalue" align="center"><B><a class="help" onmouseover="return overlib('Status When Posted:&nbsp;<%=pStatus%><br>IP Address:&nbsp; <%=ip_addr%> ');" onmouseout="return nd();">(I)</a>&nbsp;&nbsp;</b>
				</td></tr>
				<tr bgcolor='#FFFFFF'><TD colspan=4 class="<%= str_class %>"><%= server.htmlencode(replace(mydata(myfields("detail"),rowcounter),vbcrlf,vbcrlf)) %>
				<% if support_Path <> "" then %>
            <BR><B>Attachment:</b> <a href=http://<%= Site_Name %>/<%=support_path%> target=_blank class=link><%=support_path%></a>
            <% end if %>
				</td></tr>
				<!--<tr bgcolor='#FFFFFF'><TD colspan=3>IP Address: <%=mydata(myfields("ip_address"),rowcounter)%></td></tr>-->
				<!--<tr bgcolor='#FFFFFF'><TD colspan=3>Status When Posted: <%=pStatus%></td></tr>-->
				<% name = mydata(myfields("created_by"),rowcounter)
				email = mydata(myfields("email"),rowcounter)
				status = mydata(myfields("status"),rowcounter)
				store_id = mydata(myfields("store_id"),rowcounter)
				next
				end if
    if fViewOnly <> 1 then %>
				<tr bgcolor='#FFFFFF'><TD colspan=4>URL: <a href=http://<%= Site_Name %> target=_blank class=link><%= Site_Name %></a><BR>Store Id: <a href=login_as.asp?Store_Id=<%= Store_Id %> class=link><%= Store_Id %></a><BR><BR></td></tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Name</b></td>
					<td width="90%" class="inputvalue" colspan=2><input type="text" name="Name" value="Support (<%=Session("User_Name") %>)" size="30">
					<INPUT type="hidden"  name="Name_C" value="Re|String|0|150|||Name">
					<% small_help "Name" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Detail</b></td>
					<td width="90%" class="inputvalue" colspan=2><textarea name="Detail" rows=10 cols=55></textarea>
					<INPUT type="hidden"  name="Detail_C" value="Re|String|0|3000|||Detail">
					<% small_help "Detail" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Status</b></td>
					<td width="90%" class="inputvalue" colspan=2><select id="Status" name="Status" OnChange="showHideRes()">
					<option value=1>New</option>
					<option value=2>Assigned</option>
					<option value=3>Cancelled</option>
					<option value=4>Completed</option>
					<option value=5>Reopened</option>
					<option value=6>Awaiting Feedback</option>
					</select>

					<% small_help "Status" %></td>
				</tr>
			

				<!--  for assigning tasks -->
				<tr bgcolor='#FFFFFF'>
					<td colspan="4">
					<div name="divassign" id="divassign" style="position:relative; left:0px; top:0px; display:none;">
						<table cellspacing="0" cellpadding="0" border="0">
							<tr bgcolor='#FFFFFF'>					
					<td width="10%" class="inputname"><b>Assign To</b></td>
					<td width="70%" class="inputvalue" colspan=2>
					
						<select name="Assigned_to" >
						    <option value="shawn">AbuTurab (Shawn)</option>
							<option value="alec">Alec (Awais)</option>
							<option value="joe">Ali (Joe)</option>
							<option value="amer">Amer</option>
							<option value="jason">Basel (Jason)</option>
                            <option value="cheryl">Cheryl</option>
                            <option value="hadi">Hadi</option>
                            <option value="melanie">Melanie</option>
							<option value="michael">Motasim (Michael)</option>
							<option value="adam">Nauman (Adam)</option>
							<option value="rashid">Rashid</option>
							<option value="sam">Razi (Sam)</option>
							
							
					</select>

					</td>			

					<td width="10%" class="inputname"><b>Send Email to Assignee</b></td>
					<td width="10%" class="inputvalue"><input  name="send_email" value="1"  type="checkbox"></td>
					</tr>

					</table>
					</div>
					</td>
				</tr>

				<!-- assigning ends  -->
				<tr bgcolor='#FFFFFF'><td width="10%" class="inputname"><b>Delay</b></td>
					<td width="90%" class="inputvalue" colspan=3><input name=Delay value=0> hours</td></tr>
					<tr bgcolor='#FFFFFF'><td width="10%" class="inputname"><b>Send to Chatters</b></td>
					<td width="90%" colspan=3 class="inputvalue"><input  name="send_chatters" value="1"  type="checkbox"> For training purposes</td>
					</tr>
				<input type="hidden" name="Support_Id" value="<%= Support_Id %>">
				<input type="hidden" name="Store_Id" value="<%= Store_Id %>">
				<input type="hidden" name="Email" value="<%= email %>">
				<% createFoot thisRedirect, 1%>
				<% else %>
<% createFoot thisRedirect, 0%>
<% end if %>
