<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

'LOAD CURRENT SETTINGS FROM THE DATABASE

sql_select_min = "select max(OID) as theMin from store_purchases where store_id="&store_id
rs_Store.open sql_select_min,conn_store,1,1
MIN_StartOID = -1
if not rs_Store.eof then
	if rs_store("theMin")<>"" then
	MIN_StartOID = clng(rs_store("theMin"))
	end if
end if
rs_Store.close

sql_select_min = "select max(CCID) as theMin from store_customers where store_id="&store_id
rs_Store.open sql_select_min,conn_store,1,1
MIN_StartCID = -1
if not rs_Store.eof then
	if rs_store("theMin")<>"" then
		MIN_StartCID = clng(rs_store("theMin"))
	  end if
end if
rs_Store.close

sInstructions="Starting id numbers for use in your store.  If there are already id numbers used in your store they will not be renumbered.  Once you you select and use a starting id number you cannot use a number lower than what is specified."
				
sFormAction = "Store_Settings.asp"
sName = "Store_Manag"
sFormName = "Store_Manag_Starts"
sCommonName="Id Numbers"
sTitle = "Id Numbers"
sFullTitle = "General > Id Numbers"
sSubmitName = "Update_Cookies"
thisRedirect = "store_manag.asp"
sTopic = "Store_manage"
sMenu = "general"
sQuestion_Path = "advanced/store_manage.htm"
createHead thisRedirect
if Service_Type < 3 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		BRONZE Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>

				<TR bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Starting Customer ID</b></td>
				<td class="inputvalue">
						<input name="StartCID" value="<%= StartCID %>" size="10" onKeyPress="return goodchars(event,'0123456789')">
						<input type="Hidden" name="StartCID_C" value="Re|Integer|<%=MIN_StartCID %>||||Starting Customer ID">
						<% small_help "Starting Customer ID" %></td>
				</tr>
		  

		  
				<TR bgcolor='#FFFFFF'>
				<td width="1%" nowrap class="inputname"><b>Starting Order ID</b></td>
				<td class="inputvalue">
						<input name="StartOID" value="<%= StartOID %>" size="10" onKeyPress="return goodchars(event,'0123456789')">
						<input type="Hidden" name="StartOID_C" value="Re|Integer|<%=MIN_StartOID %>||||Starting Order ID">
						<% small_help "Starting Order ID" %></td>
				</tr>
		  




<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("StartOID","greaterthan=<%= MIN_StartOID %>","Please enter a starting order id greater then <%= MIN_StartOID %>");
 frmvalidator.addValidation("StartCID","greaterthan=<%= MIN_StartCID %>","Please enter a starting customer id greater then <%= MIN_StartCID %>");

</script>
<% end if %>
