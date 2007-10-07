<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
Oid = Request.QueryString("Oid")
Cid= Request.QueryString("Cid")  

sql_customers = "select email from store_customers where cid = "&cid&" and record_type = 1"
rs_Store.open sql_customers,conn_store,1,1
	email = Rs_store("email")
Rs_store.Close

sFormAction = "Return_Orders_action.asp"
sName = "Return_Orders"
sTitle = "Return Order - #"&Oid
sFullTitle = "<a href=orders.asp class=white>Orders</a> > <a href=order_details.asp?Id="&oid&" class=white>Details</a> > Return Order - #"&Oid
sMenu="orders"
thisRedirect = "return_orders.asp"
createHead thisRedirect
%>


<input type="Hidden" name="Oid" value="<%= Oid %>">
<input type="Hidden" name="Cid" value="<%= Cid %>">



				
				<TR bgcolor='#FFFFFF'>
					<td width="41%" height="25" class="inputname">Return RMA</td>
			<td width="59%" height="25" class="inputvalue">
						<input name="Return_RMA" size="60"></td>
				</tr>

				<TR bgcolor='#FFFFFF'>
			<td width="41%" height="43" class="inputname">Return Reason</td>
			<td width="59%" height="43" class="inputvalue">
						<textarea cols="32" name="Return_notes" rows="2"></textarea></td>
			</tr>
               <TR bgcolor='#FFFFFF'>
			<td height="23" width="20%" class="inputname">Send email to customer </td><td>
						<input type="checkbox" class="image" name="send_return_email" value="1" checked></td>
			</tr>

	  
				<TR bgcolor='#FFFFFF'>
			<td height="23" width="8%" class="inputname">To</td>
			<td height="23" width="92%" class="inputvalue">
						<input name="Return_Email_To" size="60" value="<%= email %>"></td>
			</tr>

				<TR bgcolor='#FFFFFF'>
			<td height="23" width="8%" class="inputname">From</td>
			<td height="23" width="92%" class="inputvalue">
						<input name="Return_Email_From" size="60" value="<%= Store_Email %>"></td>
			</tr>

				<TR bgcolor='#FFFFFF'>
			<td height="17" width="8%" class="inputname">Subject</td>
			<td height="17" width="92%" class="inputvalue">
						<input name="Return_email_subject" size="60" value="Your Return"></td>
			</tr>
	  
				<TR bgcolor='#FFFFFF'> 
			<td height="129" width="8%" class="inputname">Message Body</td>
			<td height="129" width="92%" class="inputvalue">
						<textarea cols="60" name="Return_Email_Body" rows="10"></textarea></td>
			</tr>
	  
				<TR bgcolor='#FFFFFF'>
			<td height="25" width="100%" colspan="2" align="center">
				<input name="Make_Return " type="submit" class="Buttons" value="Return Order Now"></td>
			</tr>

<% createFoot thisRedirect, 0%>
