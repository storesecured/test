<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

'RETRIVE CURRENT VALUES FROM THE DATABASE

if request.querystring("Use_Domain_Name") = "Yes" then
   sql_update = "Update store_settings set use_domain_name=1,mail_done=0 where store_Id="&Store_Id
   conn_store.Execute sql_update
   server.execute "reset_design.asp"

	'Renaming folder
	Store_Domain=replace(Store_Domain,"www.","")
	Store_Domain2=replace(Store_Domain2,"www.","")

elseif request.querystring("Use_Domain_Name") = "No" then
   sql_update = "Update store_settings set use_domain_name=0,mail_done=0 where store_Id="&Store_Id
   conn_store.Execute sql_update
	server.execute "reset_design.asp"
elseif request.querystring("delete_primary") = "yes" then
	sql_update = "Update store_settings set use_domain_name=0,store_domain='',mail_done=0 where store_Id="&Store_Id
	conn_store.Execute sql_update
	server.execute "reset_design.asp"
elseif request.querystring("delete_secondary") = "yes" then
	sql_update = "Update store_settings set store_domain2='',mail_done=0 where store_Id="&Store_Id
	conn_store.Execute sql_update

	server.execute "reset_design.asp"
end if
'********MAKING ADDITIONAL DOMAIN AS STORE ADDRESS*****************
if request.QueryString("Use_Addl_Domain_Name") ="Yes" then

	current_domain = store_domain
	new_domain = store_domain2

	sql_update = "update store_settings set use_domain_name=1, store_domain ='" & new_domain &"' , store_domain2='" & current_domain & "',mail_done=0 where store_id=" & store_id
	response.write sql_update
	conn_store.execute sql_update
	store_domain =   new_domain
	store_domain2 =current_domain
   Store_Domain=replace(store_domain,"www.","")
   Store_Domain2=replace(store_domain2,"www.","")

	Site_Name = "http://"&Store_Domain&"/"
   	server.execute "reset_design.asp"
   	response.redirect "domain.asp"

end if
'*****************END OF MAKING ADDL DOMAIN AS DEFAULT*****************

sArticleHelp="domain.htm"
sTextHelp="domainname/domainname.doc"

sFormAction = "store_settings.asp"
sName = "Domain"
sFormName = "domain"
sTitle = "Domain Name"
sFullTitle = "General > Domain Name"
sSubmitName = "Register"
thisRedirect = "domain.asp"
sTopic = "Domain"
addPicker=1
sMenu="general"
sQuestion_Path = "advanced/domain_name.htm"
createHead thisRedirect

if Service_Type=0 then
%>
	<tr bgcolor='#FFFFFF'>
		<td colspan=2>
			This feature is not available at your current level of service.<BR><BR>
			PEARL Service or higher is required.
			<a href=billing.asp class=link>Click here to upgrade now.</a>
		</td>
	</tr>
<%
else 
%>
	
	<% 
	if Store_Domain = "" then  'NO DOMAIN EXISTS
	%> 

	
	
	<tr bgcolor='#FFFFFF'>
		<td colspan=3>
			<B>Transfer a domain name you already own</B>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'>
		<td colspan=3>
			If you currently have a site at this address it will no longer be available after the transfer.
			<BR><BR>
			If you need to register the name and have not done so already we recommend <a href="http://www.anrdoezrs.net/i3102uoxuowBDFFGFCIBDCFJKGCI" target="_blank" onmouseover="window.status='http://www.godaddy.com';return true;" onmouseout="window.status=' ';return true;">www.GoDaddy.com </a>
<img src="http://www.tqlkg.com/7q105iw-ousDFHHIHEKDFEHLMIEK" width="1" height="1" border="0"/>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'>
		<td class="inputname">
			<b>Domain</b>
		</td>
		<td class="inputvalue">
			http://<input type="text" name="Store_Domain1" value="<%= Store_Domain %>" size="60" maxlength=63 onKeyPress="return goodchars(event,'0123456789abcdefghijklmnopqrstuvwxyz-.')">
			<input type="hidden" name="Store_Domain1_C" value="Op|String|0|63|.| ,%,^,&,(,),#,@,:,;,',>,<,/,=,+,{,},[,],\,~,`,$,*,_|Domain Name">
			<% small_help "Domain Transfer" %>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'>
		<td colspan=3 align=center>
			<INPUT type="submit" value="Transfer" name="Transfer">
		</td>
	</tr>

	
	

	<% end if %>
		
	<tr bgcolor='#FFFFFF'>
		<TD colspan=3>
			<B>Transfers usually take approximately 48 hours to propogate throughout the net but can take up to a week in some cases, until that time your store will
			still be available at the default address: <%= Site_Name %>.
			</B>
		</td>
	</tr>
	
	<% 
		if Store_Domain = "" then
	%>
			<tr bgcolor='#FFFFFF'>
				<td colspan=3>
					<BR>Your default store address is <a class="link" href='<%= Site_Name %>' target='_blank'><%= Site_Name %></a>
				</td>
			</tr>
		<% 
		elseif Use_Domain_Name <>0 then
		%>
			<tr bgcolor='#FFFFFF'>
				<td colspan=2>

                                        <BR>Your default store address is <a class="link" href='<%= Site_Name %>' target='_blank'><%= Site_Name %></a>
                                        <BR>Your secondary store address is <a class="link" href='<%= Site_Name2 %>' target='_blank'><%= Site_Name2 %></a>
  					<BR>To make <%= Site_Name2 %> the default store address <a class=link href="domain.asp?Use_Domain_Name=No">click here</a>
  					<BR><BR>
                                                       <table border=1 cellpadding=5 cellspacing=0 bgcolor=yellow><tr><td><B>Please wait at least 24 hours after requesting a transfer before making the name the default.  Doing so earlier could cause issues with domain name propogation.</b></td></tr></table>

				</td>
			</tr>
		<%
		else
		%>
	 		<tr bgcolor='#FFFFFF'>
	 			<td colspan=2>
					<% Set DNS = Server.CreateObject("ASPDNS.DNSLookup")
                                        Site_Ip = DNS.GetIPFromName(Store_Domain)
                                        if Site_Ip=Server_IP or Site_Ip=Server_IP2 then %>
                                             <BR>Your domain name (http://<%=Store_Domain %>) is correctly pointing at your stores server.
                                             <BR>To make <%= Store_Domain %> the default store address <a class=link href= javascript:confirm_domain_change('domain.asp?Use_Domain_Name=Yes')>click here</a>
                                             <BR><BR>
                                                       <table border=1 cellpadding=5 cellspacing=0 bgcolor=yellow><tr><td><B>Please wait at least 24 hours after requesting a transfer before making the name the default.  Doing so earlier could cause issues with domain name propogation.</b></td></tr></table>

                                        <% else %>

                                           <BR>
                                             <Table border=1 cellspacing=0 cellpadding=2><tr bgcolor='#FFFFFF'><td bgcolor=red><font color=white>
                                             WARNING...Your domain name (http://<%=Store_Domain %>) is NOT SETUP correctly at this time or your updates
                                             have not yet propogated throughout the internet, note that this can take up to
                                             48 hours.
                                             <BR>Please review the transfer instructions below if it has been more then 48 hours and you are still receiving this error.</a>
                                             </td></tr></table>
                                        <% end if 
                                        set DNS=Nothing %>
				</td>
			</tr>
			<tr bgcolor='#FFFFFF'>

			</tr>
		<!--********************** ADDITIONAL DOMAIN-->
		<% end if 
		if Store_Domain <> "" then
				if Store_Domain2 = "" or isnull(Store_Domain2) then 'NO SECONDARY DOMAIN
				%>
					<tr bgcolor='#FFFFFF'>
						<td colspan="3">
							<b>You can Transfer an Additional Domain to your Store </b>
						</td>
					</tr>
					<tr bgcolor='#FFFFFF'>
						<td colspan=3 align=center>
							<input type="button" value="Transfer Additional Domain" name="Add_Additional_Domain" OnClick=JavaScript:self.location="Additional_domain_add.asp">
						</td>
					</tr>

					<tr bgcolor='#FFFFFF'>
							
				<%
				else 'SOME ADDITIONAL DOMAIN HAS BEEN TRANSFERRED
				%>
			 		<tr bgcolor='#FFFFFF'>
			 			<td colspan=2>
                                                <% Set DNS = Server.CreateObject("ASPDNS.DNSLookup")
                                                  Site_Ip = DNS.GetIPFromName(Store_Domain2)
                                                  if Site_Ip=Server_IP or Site_Ip=Server_IP2 then %>
                                                       <BR>Your secondary domain name (http://<%=Store_Domain2 %>) is correctly pointing at your stores server.
                                                       <BR>To make <%= Store_Domain2 %> the default store address <a class=link href=javascript:confirm_domain_change('domain.asp?Use_Addl_Domain_Name=Yes')>click here</a>
                                                       <BR><BR>
                                                       <table border=1 cellpadding=5 cellspacing=0 bgcolor=yellow><tr><td><B>Please wait at least 24 hours after requesting a transfer before making the name the default.  Doing so earlier could cause issues with domain name propogation.</b></td></tr></table>

                                                  <% else %>
                                                     <BR>
                                                       <Table border=1 cellspacing=0 cellpadding=2><tr bgcolor='#FFFFFF'><td bgcolor=red><font color=white>
                                                       WARNING...Your secondary domain name (http://<%=Store_Domain2 %>) is NOT SETUP correctly at this time or your updates
                                                       have not yet propogated throughout the internet, note that this can take up to
                                                       48 hours.
                                                       <BR>Please review the transfer instructions below if it has been more then 48 hours and you are still receiving this error.</a>
                                                       </td></tr></table>
                                                  <% end if 
                                                  set DNS=Nothing %>
                                                 
                                             						</td>
					</tr>

				<%
				end if
				%>
		<!--*************************** ADDITIONAL DOMAIN-->
	<tr bgcolor='#FFFFFF'>
		<td colspan="3">
		<table>
			<tr bgcolor='#FFFFFF'>
				<td>
					<BR><BR><a href=transfer.asp class=link>Click here for transfer instructions</a>
					<% if Store_Domain<>"" then %>
						<% if Store_Domain2<>"" then %>
						<BR><BR><a href=domain.asp?delete_secondary=yes class=link>Click here to remove the secondary domain name <%= Store_Domain2 %></a>
						<% else %>
						<BR><BR><a href=domain.asp?delete_primary=yes class=link>Click here to remove the primary domain name <%= Store_Domain %></a>
						<% end if %>
					<% end if %>
					</td>
			</tr>
		</table>
		</td>
	</tr>
	
	<%
	end if ' CHECKING THE STORE DOMAIN
	%>
</td></tr>


		
<% 
end if 'CHECKING FOR THE TRIAL VERSION
%>
<% createFoot thisRedirect,2 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 
</script>
	<!-- NEW LOCATION-->
	<script language="JavaScript">


	function confirm_domain_change( domain )
	{
				
				var conf_change  = confirm("Are you sure you want to change the default store address?");
				if (conf_change == true) 
				{
					window.location =domain;
					//alert (domain);
				}

	}
	function conf_domain()
	{
		alert('hi');
	}

</script>
<!-- NEW LOCATION-->
