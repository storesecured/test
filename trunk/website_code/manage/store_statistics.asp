<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sql_select = "select * from store_statistics where store_id="&Store_Id
rs_store.open sql_select,conn_store,1,1
   gmt_offset = rs_store("gmt_offset")
   skip_hosts = rs_store("skip_hosts")
   dns_lookup = rs_store("dns_lookup")
   if dns_lookup then
      dns_checked = "checked"
   else
      dns_checked = ""
   end if
rs_store.close

sFlashHelp="statistics/statistics.htm"
sMediaHelp="statistics/statistics.wmv"
sZipHelp="statistics/statistics.zip"
sInstructions="Changes made to statistics will only effect new statistics captured."

addPicker=1
sFormAction = "Store_settings.asp"
sName = "Store_statistics"
sFormName = "statistics"
sTitle = "Statistics Settings"
sFullTitle = "Statistics > Settings"
sCommonName = "Statistics Settings"
sSubmitName = "Store_statistics"
thisRedirect = "store_statistics.asp"
sTopic="Statistics"
sMenu = "statistics"
sQuestion_Path = "advanced/statistics_settings.htm"
createHead thisRedirect
if Service_Type < 1  then %>
	<TR bgcolor=FFFFFF>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>

	</td></tr>

<% else %>

		<TR bgcolor=FFFFFF>
				<td width="15%" class="inputname"><b>GMT Offset</b></td>
				<td width="85%" class="inputvalue"><input name="GMT_Offset" type="text" value="<%= gmt_offset %>" size=3>&nbsp;hours
			<input type="hidden" name="GMT_Offset_C" value="Re|Integer|-12|13|||GMT Offset"><% small_help "GMT_Offset" %></td></tr>

		<TR bgcolor=FFFFFF>
				<td width="15%" class="inputname"><b>Skip Hosts</b></td>
				<td width="85%" class="inputvalue"><input name="Skip_Hosts" type="text" value="<%= skip_hosts %>" size=60>&nbsp;
				<input type="hidden" name="Skip_Hosts_C" value="Op|String|0|200|||Skip Hosts"><% small_help "Skip_Hosts" %></td>
			
			</tr>
		<TR bgcolor=FFFFFF>
				<td width="15%" class="inputname"><b>DNS Lookup</b></td>
				<td width="85%" class="inputvalue"><input class=image name="Dns_Lookup" type="checkbox" value="1" <%= dns_checked %>>&nbsp;
				<% small_help "Dns_Lookup" %></td>
			
			</tr>
			 

<% createFoot thisRedirect, 1%>
<% end if %>
