<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<%
Oid = request.querystring("Orders")
if Oid = "" then
   fn_error "You must select at least one order"
else

sql_store_emails = "Select Store_emails.*, html_notifications FROM Store_emails inner join store_settings on store_emails.store_id=store_settings.store_id where Store_emails.Store_id = "&Store_id
rs_Store.open sql_store_emails,conn_store,1,1
	Shipping_email_subject= Rs_store("Shipping_email_subject")
	Shipping_email_body= Rs_store("Shipping_email_body")
	html_notifications=Rs_store("html_notifications")
Rs_store.Close



sFormAction = "Ship_Orders_batch_action.asp"
sName = "Ship_Orders"
sTitle = "Orders > Ship Selected"
sMenu="orders"
thisRedirect = "ship_orders_batch.asp"
createHead thisRedirect

if Html_Notifications = 0 then
	show_editor_cookie=false
	Html_Notifications=0
else
        Html_Notifications=1
end if
dim flag
%>
<input type="Hidden" name="Orders" value="<%= Oid %>">
<input type="Hidden" name="html_notifications" value="<%= html_notifications %>">
	<tr bgcolor='FFFFFF'>
	<td colspan=2>
<table width='100%' border='0' cellspacing='1' cellpadding=2>
 <tr bgcolor="#6495ed" align="center" clas="tablehead">
      <th width="20%" height="27"><font color="#FFFFFF">Order Id</font></th>
      <th width="60%" height="27"><font color="#FFFFFF">Tracking ID</font></th>
      <th width="20%" height="27"><font color="#FFFFFF">Capture</font></th>
    </tr>
   <%

    flag=1


sArray = Split(Oid,",")
for each Oid in sArray
 Oid=Trim(Oid)
 sql_select_type = "Select Transaction_Type,Processor_id,MasterOID from Store_Purchases where oid =" & Oid &" and Store_id ="&Store_id
rs_Store.open sql_select_type,conn_store,1,1
	  Transaction_Type = Rs_store("Transaction_Type")
	  Processor_id = Rs_store("Processor_id")
	  MasterOID=Rs_store("MasterOID")
rs_Store.close
if flag=1 then
coloro="white"
else
coloro="#D4DEE5"
end if

 if (Processor_id=1 or Processor_id=2 or Processor_id=7 or Processor_id=10 or Processor_id=36) and isNull(MasterOID) then
 if Transaction_Type = "0" then
%>
 <tr align='center' height='15'><font size='2' face='verdana'><TD  align='left' height='15' bgcolor=<%=coloro%>><%= Oid%></TD></TD> <TD height='15' bgcolor=<%=coloro%> class='inputvalue' align=left><input name=Tracking_ID<%=Oid%> size=20 maxlength=50></TD><TD height='15' bgcolor=<%=coloro%> align='center'><input class='image' type='checkbox' name=capture<%=Oid%> value=<%=Oid%>></TD></tr>
 <%
elseif Transaction_Type = "2" then
 %>
 <tr align='center' height='15'><font size='2' face='verdana'><TD  align='left' height='15' bgcolor=<%=coloro%>><%= Oid%></TD></TD> <TD height='15' bgcolor=<%=coloro%> class='inputvalue' align=left><input name=Tracking_ID<%=Oid%> size=20 maxlength=50></TD><TD height='15' bgcolor=<%=coloro%> align='center'>Captured</TD></tr>
 <%
 else
 %>
 <tr align='center' height='15'><font size='2' face='verdana'><TD  align='left' height='15' bgcolor=<%=coloro%>><%= Oid%></TD></TD> <TD height='15' bgcolor=<%=coloro%> class='inputvalue' align=left><input name=Tracking_ID<%=Oid%> size=20 maxlength=50></TD><TD height='15' bgcolor=<%=coloro%> align='center'>-</TD></tr>
 <%
end if
else
  %>
 <tr align='center' height='15'><font size='2' face='verdana'><TD  align='left' height='15' bgcolor=<%=coloro%>><%= Oid%></TD></TD> <TD height='15' bgcolor=<%=coloro%> class='inputvalue' align=left><input name=Tracking_ID<%=Oid%> size=20 maxlength=50></TD><TD height='15' bgcolor=<%=coloro%> align='center'>-</TD></tr>
 <%
end if

flag=-flag
Next
  %>
  </table>
</div>
</td>
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
			<td width="20%" height="43" class="inputname">Special notes</td>
			<td width="80%" height="43" class="inputvalue">
						<textarea cols="50" name="Shippment_notes" rows="2"></textarea></td>
			</tr>


				<tr bgcolor='FFFFFF'>
			<td height="23" width="100%" colspan="2" class="inputname">Send email to customer
						<input type="checkbox" class="image" name="send_shipment_email" value="1" checked></td>
			</tr>

				<tr bgcolor='FFFFFF'>
			<td height="23" width="8%" class="inputname">From</td>
			<td height="23" width="92%" class="inputvalue">
						<input name="Shipping_Email_From" size="40" value="<%= Store_Email %>"></td>
			</tr>
	  
				<tr bgcolor='FFFFFF'>
			<td height="17" width="8%" class="inputname">Subject</td>
			<td height="17" width="92%" class="inputvalue">
						<input name="Shipping_email_subject" size="60" value="<%= Shipping_Email_Subject %>"></td>
			</tr>
	  
				<tr bgcolor='FFFFFF'> 
			<td colspan=2 class="inputname">Message Body</td></tr>
			<tr bgcolor='FFFFFF'><td colspan=2 width="92%" class="inputvalue">

                        <%= create_editor ("Shipping_Email_Body",Shipping_email_body,"") %>
								</td></tr>
	  
				<tr bgcolor='FFFFFF'>
			<td height="25" width="100%" colspan="2" align="center">
				<input name="Make_Shipment " type="submit" class="Buttons" value="Ship Order Now"></td>
			</tr>





<% createFoot thisRedirect, 0%><% end if %>

