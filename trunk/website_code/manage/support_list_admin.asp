<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
if Session("Super_User") <> 1 then
     Response.redirect "noaccess.asp"
end if

if request.querystring("Support_Id") <> "" and request.querystring("Status") <> "" then
	sql_update = "update sys_support WITH (ROWLOCK) set Status="&request.querystring("Status")&" where support_Id="&request.querystring("Support_Id")
	conn_store.Execute sql_update

	Detail=""
    Name="Support"
   	IP_Address= Request.ServerVariables("REMOTE_ADDR")
	'sql_insert = "Insert_Support_Request1 "&request.querystring("Support_Id")&",'"&Detail&"','"&Name&"','"&IP_Address&"',"&request.querystring("Status")
	'sql_insert = "insert into Sys_Support_Detail (Support_Id,Detail,Created_By,IP_Address,poststatus) values ("&request.querystring("Support_Id")&",'"&Detail&"','"&Name&"','"&IP_Address&"',"&request.querystring("Status")&")"
	
	sql_insert="New_Support_Details1 "&request.querystring("Support_Id")&",'"&Detail&"','"&Name&"','"&IP_Address&"',"&request.querystring("Status")
'	Response.write sql_insert
'	Response.end
	conn_store.Execute sql_insert
'Response.write "<br> sql_insert " & sql_insert
'Response.end

end if

sql_select = "select top 50 sys_support.sys_created, sys_support.assigned_to,subject, support_id,status,sys_support.store_id,service_type,site_name,service_type,trial_version from Sys_Support WITH (NOLOCK) inner join store_settings WITH (NOLOCK) on sys_support.store_id=store_settings.store_id where status <> 3 and status <> 4 and status<>6 order by status,sys_support.sys_created"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

sFormAction = "support_list_admin.asp"
sTitle = "Admin Support Requests"
thisRedirect = "support_list_admin.asp"
sMenu = "account"
createHead thisRedirect
%>

  
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="13">
			<table width="100%" border="1" cellspacing="0" cellpadding=2 class="list">
			<tr bgcolor=#DDDDDD>
			<td><b>Date</b></td>
			<td><b>Subject</b></td>
			<td><b>S</b></td>
			<td><b>Store</b></td>
			<td><b>L</b></td>
			<td><b>&nbsp;</b></td>

			</tr>

			<% str_class=1
				if noRecords = 0 then
				FOR rowcounter= 0 TO myfields("rowcount")

				if str_class=1 then
			  str_class=0
				 else
				str_class=1
				 end if %>
					<tr class="<%= str_class %>">
					<td><%= replace(formatdatetime(mydata(myfields("sys_created"),rowcounter),2),"/2004","") %></td>
					<td><%= mydata(myfields("subject"),rowcounter) %></td>

					<%
					Support_Id = mydata(myfields("support_id"),rowcounter)
					iStatus = mydata(myfields("status"),rowcounter)
					if iStatus = 1 then
						sStatus = "N"
					elseif iStatus = 2 then
						sStatus = "A"
					elseif iStatus = 3 then
						sStatus = "CN"
					elseif iStatus = 4 then
						sStatus = "CL"
					elseif iStatus = 5 then
						sStatus = "RE"
					elseif iStatus = 6 then
						sStatus = "F"
					end if 
					
					
					if mydata(myfields("status"),rowcounter)="2" then
						iAdmin= mydata(myfields("assigned_to"),rowcounter)
						if iAdmin=1 then
							nAdmin="Melanie"
						elseif iAdmin=2 then
							nAdmin="Abbas"
						elseif iAdmin=3 then
							nAdmin="Chris"
						elseif iAdmin=4 then
							nAdmin="Ganesh"
						elseif iAdmin=5 then
							nAdmin="Cheriyan"
						elseif iAdmin=6 then
							nAdmin="Ashwini"
						elseif iAdmin=7 then
							nAdmin="Sahil"
						else
							nAdmin=""
						end if
					else
							nAdmin=""
					end if				
				
								
					
					%>
					<td><%= sStatus %></td>
					<td><a class=link target=_blank href=http://<%= mydata(myfields("site_name"),rowcounter) %>><%= mydata(myfields("store_id"),rowcounter) %>-<%= left(replace(mydata(myfields("site_name"),rowcounter),".easystorecreator.com",""),7) %></a></td>
					<% if mydata(myfields("trial_version"),rowcounter) then
                                              Trial="-Trial"
                                        else
                                              Trial=""
					end if %>
                                        <td><%= mydata(myfields("service_type"),rowcounter) %><%= Trial %></td>
					<% if iStatus = 4 or iStatus = 3 then %>
					<td><a class=link href=support_edit_admin.asp?Support_Id=<%= Support_Id %>&ViewOnly=1>View</a> /
					<a class=link href=support_list_admin.asp?Support_Id=<%= Support_Id %>&Status=5>Reopen</a></td>

					<% else %>
					<td><a class=link href=support_edit_admin.asp?Support_Id=<%= Support_Id %>>E</a>/<a class=link href=support_list_admin.asp?Support_Id=<%= Support_Id %>&Status=6>F</a>/<a class=link href=support_list_admin.asp?Support_Id=<%= Support_Id %>&Status=3>CN</a>
					<!--/<a class=link href=support_list_admin.asp?Support_Id=<%= Support_Id %>&Status=2>A</a>-->
					/<a class=link href=support_list_admin.asp?Support_Id=<%= Support_Id %>&Status=4>CL</a></td>

						<% if mydata(myfields("status"),rowcounter) = 2 then %>
							<td class="inputvalue" align="center"><B><a class="help" onmouseover="return overlib('Assigned to:&nbsp;<%=nAdmin%>');" onmouseout="return nd();">(I)</a>&nbsp;&nbsp;</b></td>
						<% end if %>		
					<% end if %>
				</tr>
				<% Next
				end if %>
			</table>
		</td>
	</tr>
<% createFoot thisRedirect, 0%>

