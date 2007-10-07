<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%
Oid = Request.QueryString("Oid")
Cid= Request.QueryString("Cid")  
sql_store_emails = "Select Store_emails.*, html_notifications FROM Store_emails inner join store_settings on store_emails.store_id=store_settings.store_id where Store_emails.Store_id = "&Store_id
rs_Store.open sql_store_emails,conn_store,1,1
	Shipping_email_subject= Rs_store("Shipping_email_subject")
	Shipping_email_body= Rs_store("Shipping_email_body")
	html_notifications=Rs_store("html_notifications")
Rs_store.Close

sql_select_type = "Select Transaction_Type,Processor_id from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id
rs_Store.open sql_select_type,conn_store,1,1
     Transaction_Type = Rs_store("Transaction_Type")
     Processor_id = Rs_store("Processor_id")
rs_Store.close

sql_customers = "select shipfirstname,shiplastname,shipemail,user_id,password from store_purchases p with (NOLOCK) inner join store_customers c WITH (NOLOCK) on c.cid=p.cid and c.store_id=p.store_id where p.store_id="&store_id&" and p.oid = "&oid&" and record_type = 0"
rs_Store.open sql_customers,conn_store,1,1
	email = Rs_store("shipemail")
	last_name = Rs_store("shiplastname")
	first_name = Rs_store("shipfirstname")
	username = Rs_store("user_id")
	password = Rs_store("password")
Rs_store.Close

'Shipping_email_body = Replace(Shipping_email_body,"%FIRSTNAME%",first_name)
'Shipping_email_body = Replace(Shipping_email_body,"%LASTNAME%",last_name)
'Shipping_email_body = Replace(Shipping_email_body,"%LOGIN%",username)
'Shipping_email_body = Replace(Shipping_email_body,"%PASSWORD%",password)

sFormAction = "Ship_Orders_action.asp"
sName = "Ship_Orders"
sMenu="reports"
sTitle = "Ship Order - #"&Oid
sFullTitle = "<a href=orders.asp class=white>Orders</a> > <a href=order_details.asp?Id="&oid&" class=white>Details</a> > Ship Order - #"&Oid
thisRedirect = "ship_orders.asp"
createHead thisRedirect

if Html_Notifications = 0 then
	show_editor_cookie=false
	Html_Notifications=0
else
        Html_Notifications=1
end if
%>


<input type="Hidden" name="Oid" value="<%= Oid %>">
<input type="Hidden" name="Cid" value="<%= Cid %>">
<input type="Hidden" name="last_name" value="<%= last_name %>">
<input type="Hidden" name="first_name" value="<%= first_name %>">
<input type="Hidden" name="username" value="<%= username %>">
<input type="Hidden" name="password" value="<%= password %>">
<input type="Hidden" name="html_notifications" value="<%= html_notifications %>">



				
				<tr bgcolor='#FFFFFF'>
					<td width="20%" height="25" class="inputname">Tracking ID for Shipment</td>
			<td width="80%" height="25" class="inputvalue">
						<input name="Tracking_ID" size="60" maxlength=50></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="20%" height="25" class="inputname">Shipment Company</td>
			<td width="80%" height="25" class="inputvalue">
						<select name="Shipping_Company">
            <option value=>Please Select</option>
            <option value=Other>Other</option>
            <option value=USPS>USPS</option>
            <option value=UPS>UPS</option>
            <option value=Fedex>Fedex</option>
            <option value=DHL>DHL</option>
            <option value=Airborne>Airborne</option>
            <option value=CanadaPost>CanadaPost</option>
            </select></td>
				</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td width="20%" height="36" class="inputname">Full Shipping</td>
			<td width="80%" height="36" class="inputvalue"><input class="image" CHECKED name="ShippedPr" type="radio" value="yes">Yes, shipping whole order&nbsp;<br>
				<input name="ShippedPr" class="image" type="radio" value="Partial">No, Partial Shipment</td>
			</tr>

				<tr bgcolor='#FFFFFF'>
			<td width="20%" height="43" class="inputname">Special notes</td>
			<td width="80%" height="43" class="inputvalue">
						<textarea cols="50" name="Shippment_notes" rows="2"></textarea></td>
			</tr>


				<tr bgcolor='#FFFFFF'>
			<td height="23" width="20%" class="inputname">Send email to customer</td><td>
						<input type="checkbox" class="image" name="send_shipment_email" value="1" checked></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td height="23" width="20%" class="inputname">To</td>
			<td height="23" width="80%" class="inputvalue">
						<input name="Shipping_Email_To" size="60" value="<%= email %>"></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td height="23" width="20%" class="inputname">From</td>
			<td height="23" width="80%" class="inputvalue">
						<input name="Shipping_Email_From" size="60" value="<%= Store_Email %>"></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td height="17" width="20%" class="inputname">Subject</td>
			<td height="17" width="80%" class="inputvalue">
						<input name="Shipping_email_subject" size="60" value="<%= Shipping_Email_Subject %>"></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'> 
			<td height="129" width="100%" class="inputname" colspan=2>Message Body<BR>
			<%= create_editor ("Shipping_email_body",Shipping_email_body,"[""First Name"",""%FIRSTNAME%""],[""Last Name"",""%LASTNAME%""],[""Login"",""%LOGIN%""],[""Password"",""%PASSWORD%""],[""Order Id"",""%ORDER%""],[""Tracking #"",""%TRACKING%""]") %>
			</td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
			<td height="25" width="100%" colspan="2" align="center">
				<input name="Make_Shipment " type="submit" class="Buttons" value="Ship Order Now"></td>
			</tr>
                        
			<tr bgcolor='#FFFFFF'>
					<td height="19" colspan=2>
						<hr width="600" align="left">
							<b>
							<a class="link" href="ups_print_labels.asp?Orders=<%= oid %>">Shipping Labels</a></b>&nbsp;

					</td>
				</tr>

<% createFoot thisRedirect, 0%>
