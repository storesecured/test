<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

if Session("Super_User") <> 1 then
    response.Redirect "noaccess.asp"
end if

if request.Form<>"" then
    additional_storage = request.form("additional_storage")
	bandwidth_subscription = request.form("bandwidth_subscription")
	email_offsite = request.form("email_offsite")
	service_type = request.form("service_type")
	trial_version = request.form("trial_version")
	overdue_payment = request.form("overdue_payment")
	if trial_version <> "" then
		trial_version = 1
	else
		trial_version = 0
	end if
	if email_offsite <> "" then
		email_offsite = 1
	else
		email_offsite = 0
	end if
	next_billing_date = request.form("next_billing_date")
	
	sql_update = "update store_settings set additional_storage="&additional_storage&", bandwidth_subscription="&_
	    bandwidth_subscription&", email_offsite="&email_offsite&", service_type="&service_type&", trial_version="&_
	    trial_version&", overdue_payment="&overdue_payment&" where store_id="&store_id
	conn_store.execute sql_update
	
	sql_update = "update sys_billing set next_billing_date='"&next_billing_date&"' where store_id="&store_id
	conn_store.execute sql_update
end if

sql_Select = "Select additional_storage,bandwidth_subscription,content_27_done,content_28_done,email_offsite,ftp_done,last_login,mail_done,service_type,trial_version,overdue_payment From Store_settings WITH (NOLOCK) where Store_id = "&Store_id
rs_Store.open sql_Select,conn_store,1,1
rs_Store.MoveFirst 
	additional_storage = Rs_store("additional_storage")
	bandwidth_subscription = Rs_store("bandwidth_subscription")
	content_27_done = Rs_store("content_27_done")
	content_28_done = Rs_store("content_28_done")
	email_offsite = Rs_store("email_offsite")
	ftp_done = Rs_store("ftp_done")
	last_login = Rs_store("last_login")
	mail_done = Rs_store("mail_done")
	service_type = Rs_store("service_type")
	trial_version = Rs_store("trial_version")
	overdue_payment = Rs_store("overdue_payment")
rs_Store.Close

sql_Select = "Select next_billing_date From sys_billing WITH (NOLOCK) where Store_id = "&Store_id
rs_Store.open sql_Select,conn_store,1,1
rs_Store.MoveFirst 
	next_billing_date = Rs_store("next_billing_date")
rs_Store.Close

if trial_version <> 0 then
	trial_version_checked = "checked"
else
	trial_version_checked = ""
end if
if email_offsite <> 0 then
	email_offsite_checked = "checked"
else
	email_offsite_checked = ""
end if

if service_type=0 then
    service_type_name="Free"
elseif service_type=1 then
    service_type_name="Pearl"
elseif service_type=3 then
    service_type_name="Bronze"
elseif service_type=5 then
    service_type_name="Silver"
elseif service_type=7 then
    service_type_name="Gold"
elseif service_type=9 then
    service_type_name="Platinum"
elseif service_type=11 then
    service_type_name="Unlimited"
else
    service_type_name="Other"
end if 
   
sInstructions="This information is for modification only by admins."
sFormAction = "admin_settings.asp"
sName = "Admin Settings"
sCommonName="Admin Settings"
sFormName = "Admin Settings"
sTitle = "Admin Settings"
sFullTitle = "Admin > Settings"
sSubmitName = "Update"
thisRedirect = "admin_settings.asp"
addPicker=1
sMenu = "admin"

createHead thisRedirect
%>
		
		
		<tr bgcolor=#FFFFFF>
		<td width="24%" height="23" class="inputname">Additional Disk Storage</td>
		<td width="76%" height="23" class="inputvalue">
				<input type="text" name="additional_storage" value="<%= additional_storage %>" size="2" maxlength=2>* X 100MB
				<input type="hidden" name="additional_storage_C" value="Re|Integer|0|99|||Additional Storage">
				<% small_help "additional_storage" %></td>
		</tr>
		<tr bgcolor=#FFFFFF>
		<td width="24%" height="23" class="inputname">Bandwidth Subscription</td>
		<td width="76%" height="23" class="inputvalue">
				<input type="text" name="bandwidth_subscription" value="<%= bandwidth_subscription %>" size="2" maxlength=2>*
				<input type="hidden" name="bandwidth_subscription_C" value="Re|Integer|0|99|||bandwidth_subscription"> x 5GB
				<% small_help "bandwidth_subscription" %></td>
		</tr>
		<tr bgcolor=#FFFFFF><td colspan=3>
		Bandwidth subscriptions are charged $15 per 5GB subscription per month.  
		</td></tr>
		<tr bgcolor=#FFFFFF>
		<td width="24%" height="23" class="inputname">Days Overdue</td>
		<td width="76%" height="23" class="inputvalue">
				<input type="text" name="overdue_payment" value="<%= overdue_payment %>" size="2" maxlength=2>*
				<input type="hidden" name="overdue_payment_C" value="Re|Integer|0|99|||overdue_payment">
				<% small_help "overdue_payment" %></td>
		</tr>
		<tr bgcolor=#FFFFFF><td colspan=3>
		Anything over 0 on the overdue setting will not allow the customer to login to their store.  If a store is overdue you can 
		manually set it to 0 to allow the user to login however at midnight the next night its going to show as overdue again unless 
		a payment is recorded.
		<BR>Use 100 Days Overdue for a store on hold
		</td></tr>
		<tr bgcolor=#FFFFFF>
		<td width="24%" height="23" class="inputname">Next Billing Date</td>
		<td width="76%" height="23" class="inputvalue">
				<input type="text" name="next_billing_date" value="<%= next_billing_date %>">
				<input type="hidden" name="next_billing_date_C" value="Re|date|||||Next Billing Date">
				<% small_help "next_billing_date" %></td>
		</tr>

		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Email Hosted Offsite</td>
		<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="email_offsite" value="-1" <%= email_offsite_checked %>>&nbsp;
			<% small_help "email_offsite" %></td>
		</tr>
		<tr bgcolor=#FFFFFF><td colspan=3>
		To totally remove a users email from our control check this box, remove the email domain from the email program and update the mx records for the customer. 
		</td></tr>
		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Trial Version</td>
		<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="trial_version" value="-1" <%= trial_version_checked %>>&nbsp;
			<% small_help "trial_version" %></td>
		</tr>
		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Service Type</td>
		<td width="70%" class="inputvalue"><select name=service_type>
		<option value=<%= service_type%>><%= service_type_name%></option>
		<option value=0>Free</option>
		<option value=1>Pearl</option>
		<option value=3>Bronze</option>
		<option value=5>Silver</option>
		<option value=7>Gold</option>
		<option value=9>Platinum</option>
		<option value=11>Unlimited</option>
		
		</select>&nbsp;
			<% small_help "service_type" %></td>
		</tr>
		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Server 1 Setup</td>
		<td width="70%" class="inputvalue"><%= content_27_done %>&nbsp;
			<% small_help "content_27_done" %></td>
		</tr>
		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Server 2 Setup</td>
		<td width="70%" class="inputvalue"><%= content_28_done %>&nbsp;
			<% small_help "content_28_done" %></td>
		</tr>
		
		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Mail Done</td>
		<td width="70%" class="inputvalue"><%= mail_done %>&nbsp;
			<% small_help "mail_done" %></td>
		</tr>
		<TR bgcolor='#FFFFFF'>
		<td width="30%" class="inputname">Last Login</td>
		<td width="70%" class="inputvalue"><%= last_login %>&nbsp;
			<% small_help "last_login" %></td>
		</tr>
		
		

<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

 }
</script>

