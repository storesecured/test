<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/emails.asp"-->

<% 
'RETRIEVE CURRENT SETTINGS FROM THE DATABASE
sql_emails = "SELECT store_emails.*, html_notifications FROM Store_emails inner join store_settings on store_emails.store_id=store_settings.store_id WHERE store_settings.Store_id="&Store_id
rs_store.open sql_emails,conn_store,1,1

	if Rs_store("Registration_enable") <>	0	then
		Registration_enable_checked = "Checked"
	else
		Registration_enable_checked = ""
	end if

	if Rs_store("Registration_user_psw") <>  0  then
		Registration_user_psw_checked = "Checked"
	else
		Registration_user_psw_checked = ""
	end if

	if Rs_store("order_notification_invoice") <>  0  then
		order_notification_invoice_checked = "Checked"
	else
		order_notification_invoice_checked = ""
	end if

	if Rs_store("order_notification_to_admin_enable") <>	0	then
		order_notification_to_admin_enable_checked = "Checked"
	else
		order_notification_to_admin_enable_checked = ""
	end if

	if Rs_store("order_notification_enable") <>	0	then
		order_notification_enable_checked = "Checked"
	else
		order_notification_enable_checked = ""
	end if
		
	Registration_subject = Rs_store("Registration_subject")
	Registration_body = Rs_store("Registration_body")
	Registration_sent_to = Rs_store("Registration_sent_to")
	order_notification_subject = Rs_store("order_notification_subject")		
	order_notification_body = Rs_store("order_notification_body")
	order_notification_sent_to = Rs_store("order_notification_sent_to")
	Shipping_email_subject = Rs_store("Shipping_email_subject")
	Shipping_email_body = Rs_store("Shipping_email_body") 
	Gift_email_Subject = Rs_store("Gift_email_Subject")
	Gift_email_body = Rs_store("Gift_email_body")
	Gift_buyer_email_body = Rs_store("Gift_buyer_email_body")
	Newsletter_notification = Rs_store("Newsletter_notification")
	Newsletter_notification_subject = Rs_store("Newsletter_notification_subject")	
	pin_email_Subject = Rs_store("pin_email_Subject") 
	pin_email_body = Rs_store("pin_email_body") 
	pin_buyer_email_body = Rs_store("pin_buyer_email_body") 
	Html_Notifications=rs_store("Html_Notifications")
rs_store.close
if isNull(Newsletter_notification) then
   Newsletter_notification=""
end if
			
sFormAction = "Store_Settings.asp"
sName = ""
sFormName = "Emails"
sSubmitName = "Update"
thisRedirect = "emails.asp"
sTopic="Emails"
sMenu = "general"
createHead thisRedirect

if Html_Notifications <> 0 then
	checked_Html_Notifications = "checked"
else
	checked_Html_Notifications = ""
	show_editor_cookie=false
end if

if Service_Type < 3 then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=3>
		This feature is not available at your current level of service.<BR><BR>
		BRONZE Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect,0 %>
<% else	%>

	<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Send Notifications as HTML</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Html_Notifications" value="-1" <%= checked_Html_Notifications %>>&nbsp;
				<% small_help "HTML Notifications" %></td>
				</tr>

						  <tr align='center' bgcolor='#0069B2' class='white'>
					<td colspan=3><font class=white><b>Registration Email</b></font></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
					<td colspan=3 class=instructions>The registration email can contain the following special identifiers in the subject or body:<BR>%FIRSTNAME% %LASTNAME% %LOGIN% %PASSWORD%</td>
			</tr>
	 
				<tr bgcolor='#FFFFFF'>

					<td width="45%" class="inputname">Send Registration Email</td>
					<td width="55%" class="inputvalue">
						<input class="image" type="checkbox" name="Registration_enable" value="-1" <%= Registration_enable_checked %>>
						<% small_help "Send Registration Email" %></td>
			</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="40%" class="inputname">Include Username and Password</td>
					<td width="60%" class="inputvalue">
						<input class="image"type="checkbox" name="Registration_user_psw" value="-1" <%= Registration_user_psw_checked %>>
						<% small_help "Include Username and Password" %></td>

			</tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Subject</td>
					<td width="1%" class="inputvalue">
						<input type="text" name="Registration_subject" value="<%= Registration_subject %>" size="50" maxlength=250>
					<input type="hidden" name="Registration_subject_C" value="Op|String|0|250|||Registration Subject">
										<% small_help "Registration Subject" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname" colspan=2>Body<BR>
					<%= create_editor ("Registration_body",Registration_body,"[""First Name"",""%FIRSTNAME%""],[""Last Name"",""%LASTNAME%""],[""Login"",""%LOGIN%""],[""Password"",""%PASSWORD%""]") %>
					<input type="hidden" name="Registration_body_C" value="Op|String|0|3000|||Registration Body">
										<% small_help "Registration Body" %></td>
			</tr>
			
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Send Copy to</td>
					<td width="1%" class="inputvalue">
						<textarea rows="1" name="Registration_sent_to" cols="55"><%= Registration_sent_to %></textarea>
					<input type="hidden" name="Registration_sent_to_C" value="Op|String|0|250|@,.||Registration CC">
										<% small_help "Registration Send Copy to" %></td>
			</tr>
	 
				
	 
				<tr align='center' bgcolor='#0069B2' class='white'>
					<td colspan=3><font class=white><b>Order Notification Email</b></font></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
					<td colspan=3 class=instructions>The order notification email can contain the following special identifiers in the subject or body:<BR>%FIRSTNAME% %LASTNAME% %LOGIN% %PASSWORD% %ORDER% %INVOICE_TEXT%</td>
			</tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="40%" class="inputname">Send to Store Admin</td>
					<td width="60%" class="inputvalue">
						<input class="image"type="checkbox" name="order_notification_to_admin_enable" value="-1" <%= order_notification_to_admin_enable_checked %>>
						(Admin Email=<%= Store_Email %>) <% small_help "Send a copy of Order email to Store Admin" %></td>

			 </tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Send to Customer</td>
					<td width="1%" class="inputvalue">
						<input class="image"type="checkbox" name="order_notification_enable" value="-1" <%= order_notification_enable_checked %>>
						<% small_help "Send a copy of Order email to customer" %></td>

			</tr>
			
				<tr bgcolor='#FFFFFF'>
					<td width="40%" class="inputname">Include Invoice</td>
					<td width="60%" class="inputvalue">
						<input class="image"type="checkbox" name="order_notification_invoice" value="-1" <%= order_notification_invoice_checked %>>
						<% small_help "Include Invoice" %></td>


			 </tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Subject</td>
					<td width="1%" class="inputvalue">
						<input type="text" name="order_notification_subject" value="<%= order_notification_subject %>" size="50" maxlength=250>
					<input type="hidden" name="order_notification_subject_C" value="Op|String|0|250|||Order Subject">
										<% small_help "Order Subject" %></td>
			</tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="100%" class="inputname" colspan=2>Body<BR>
					<%= create_editor ("order_notification_body",order_notification_body,"[""First Name"",""%FIRSTNAME%""],[""Last Name"",""%LASTNAME%""],[""Login"",""%LOGIN%""],[""Password"",""%PASSWORD%""],[""Order"",""%ORDER%""],[""Invoice Text"",""%INVOICE_TEXT%""]") %>
					<input type="hidden" name="order_notification_body_C" value="Op|String|0|3000|||Order Body">
										<% small_help "Order Body" %></td>
			</tr>
			</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Send Copy To</td>
					<td width="1%" class="inputvalue">
						<textarea rows="1" name="order_notification_sent_to" cols="55"><%= order_notification_sent_to %></textarea>
					<input type="hidden" name="order_notification_sent_to_C" value="Op|String|0|250|@,.||Order CC">
										<% small_help "Order Send Copy to" %></td>
			</tr>
				
				<tr align='center' bgcolor='#0069B2' class='white'>
					<td colspan=3><font class=white><b>Shipping Email</b> (admin option at shipment time)</font></td>
			</tr>
			<tr bgcolor='#FFFFFF'>
					<td colspan=3 class=instructions>The shipping email can contain the following special identifiers in the subject or body:<BR>%FIRSTNAME% %LASTNAME% %LOGIN% %PASSWORD% %ORDER% %TRACKING% %SHIPCOMPANY%</td>
			</tr>
	 
			
	 
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Subject</td>
					<td width="1%" class="inputvalue">

						<input type="text" name="Shipping_email_subject" value="<%= Shipping_email_subject %>" size="50" maxlength=250>
					<input type="hidden" name="Shipping_email_subject_C" value="Op|String|0|250|||Shipping Subject">
										<% small_help "Shipping Subject" %></td>
			</tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname" colspan=2>Body<BR>
					<%= create_editor ("Shipping_email_body",Shipping_email_body,"[""First Name"",""%FIRSTNAME%""],[""Last Name"",""%LASTNAME%""],[""Login"",""%LOGIN%""],[""Password"",""%PASSWORD%""],[""Order Id"",""%ORDER%""],[""Tracking #"",""%TRACKING%""]") %>
					<input type="hidden" name="Shipping_email_body_C" value="Op|String|0|3000|||Shipping Body">
										<% small_help "Shipping Body" %></td>
			</tr>
	
			
	
				<tr align='center' bgcolor='#0069B2' class='white'>
					<td colspan=3><font class=white><b>Gift Certificate Email</b></font></td>
			</tr>
           <tr bgcolor='#FFFFFF'>
					<td colspan=3 class=instructions>The gift certificate email can contain the following special identifiers in the body:<BR>%AMOUNT% %GIFT_CODE%</td>
			</tr>
	 
				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Subject</td>
					<td width="1%" class="inputvalue">

						<input type="text" name="Gift_email_Subject" value="<%= Gift_email_Subject %>" size="50" maxlength=250>
					<input type="hidden" name="Gift_email_Subject_C" value="Op|String|0|250|||Gift Subject">
										<% small_help "Gift Certificate Subject" %></td>
			</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname" colspan=2>Body<BR>
                       <%= create_editor ("Gift_buyer_email_body",Gift_buyer_email_body,"[""Certificate Amount"",""%AMOUNT%""],[""Certificate ID"",""%GIFT_CODE%""]") %>
				       <input type="hidden" name="Gift_buyer_email_body_C" value="Op|String|0|3000|||Buyer Gift Body">
					<% small_help "Gift Certificate Body for Buyer" %></td>
			</tr>
			
			
	
				<tr align='center' bgcolor='#0069B2' class='white'>
					<td colspan=3><font class=white><b>Newsletter Signup Email</b></td>
			</tr>
               <tr bgcolor='#FFFFFF'>
					<td width="1%" class="inputname">Subject</td>
					<td width="1%" class="inputvalue">

						<input type="text" name="Newsletter_notification_subject" value="<%= newsletter_notification_subject %>" size="50" maxlength=250>
					<input type="hidden" name="Newsletter_notification_Subject_C" value="Op|String|0|250|||Newsletter Subject">
										<% small_help "Newsletter Notification Subject" %></td>
			</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="100%" class="inputname" colspan=2>Body<BR>
					<%= create_editor ("Newsletter_notification",Newsletter_notification,"") %>
				       <input type="hidden" name="Newsletter_notification_C" value="Op|String|0|3000|||Newsletter Notification">
  		                       <% small_help "Newsletter notification" %></td>
			</tr>
			
			
	
	<tr align='center' bgcolor='#0069B2' class='white'>
		<td colspan=3><font class=white><b>Pin Email</b></font></td>
	</tr>
	<tr bgcolor='#FFFFFF'>
		<td width="1%" class="inputname">Subject</td>
		<td width="1%" class="inputvalue">
			<input type="text" name="pin_email_Subject" value="<%= pin_email_Subject %>" size="50" maxlength=250>
			<input type="hidden" name="pin_email_Subject_C" value="Op|String|0|250|||pin Subject">
			<% small_help "Pin_Subject" %>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'>
			<td width="1%" class="inputname" colspan=2>Body<BR>
			<%= create_editor ("pin_buyer_email_body",pin_buyer_email_body,"") %>
                       <input type="hidden" name="pin_buyer_email_body_C" value="Op|String|0|3000|||Buyer pin Body">
			<% small_help "pin_body" %>
		</td>
	</tr>	
	
			
	
<% createFoot thisRedirect,1 %>
<% end if %>
