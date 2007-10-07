<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

if Session("Super_User") <> 1 then
     Response.redirect "noaccess.asp"
end if

noRecords=1

sFormAction = "domain_check.asp"
sTitle = "Domain Check"
thisRedirect = "domain_check.asp"
sMenu = "none"
createHead thisRedirect

if request("Domain") <> "" then
   sql_select = "select store_id,overdue_payment,service_type,store_domain,store_domain2,site_name,store_cancel from store_settings WITH (NOLOCK) where store_domain like '%"&trim(request("Domain"))&"%' or store_domain2 like '%"&request("Domain")&"%' or site_name like '%"&request("Domain")&"%'"
   noRecords=2
   set myfields=server.createobject("scripting.dictionary")
   Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
end if  %>
					<% str_class=1
				if noRecords = 0 then %>
				<tr bgcolor='#FFFFFF'>
					<td>Store</td><td>Overdue</td><td>Service</td><TD>Site Name</td><td>Domain</td><Td>Domain2</td><td>Cancel</td></tr>

				<% FOR rowcounter= 0 TO myfields("rowcount")

				if str_class=1 then
			  str_class=0
				 else
				str_class=1
				 end if %>


					<tr bgcolor='#FFFFFF'><td><a href=login_as.asp?Store_Id=<%= mydata(myfields("store_id"),rowcounter) %> class=link><%= mydata(myfields("store_id"),rowcounter) %></a></td>
					<td align=center><%= mydata(myfields("overdue_payment"),rowcounter) %></td>
					<td align=center><%= mydata(myfields("service_type"),rowcounter) %></td>
					<td><a href=http://<%= mydata(myfields("site_name"),rowcounter) %> target=_blank><%= mydata(myfields("site_name"),rowcounter) %></a></td>
					<td><%= mydata(myfields("store_domain"),rowcounter) %></td>
					<td><%= mydata(myfields("store_domain2"),rowcounter) %></td>
          <td><%= mydata(myfields("store_cancel"),rowcounter) %></td>
				</tr>
				<% Next
				elseif noRecords = 2 then %>
               <tr bgcolor='#FFFFFF'><TD>No records found</td></tr>
        <% end if%>


<tr bgcolor='#FFFFFF'><td colspan=7><table><tr bgcolor='#FFFFFF'>
					<td width="24%" height="23" class="inputname"><B>Domain Name</B></td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Domain" value="" size="60" maxlength=200>
							</td>
					</tr></table></td></tr>

<% createFoot thisRedirect,1 %>

