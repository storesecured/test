<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="help/modify_customer.asp"-->
<% 


ccid = int(Request.QueryString("id"))
sql_select_cust =  "SELECT Protected_page_access,CCID, Tax_Exempt, Tax_ID,User_id, Cid, Password, Budget_left, Budget_Orig, Spam, Reward_Total, Reward_Left FROM Store_Customers WHERE cCid="&ccid&" AND Record_type =	0 and store_id="&Store_Id

if ccid="" or ccid=0 then
   cid=Request.QueryString("cid")
   sql_select_cust =  "SELECT Protected_page_access,CCID, Tax_Exempt, Tax_ID,User_id, Cid, Password, Budget_left, Budget_Orig, Spam, Reward_Total, Reward_Left FROM Store_Customers WHERE Cid="&cid&" AND Record_type =	0 and store_id="&Store_Id
end if
'LOAD USER DATA, USING RECORD_TYPE PARAM:  
'		0=BILLING ADDRESS
'		1=SHIPPING ADDRESS 1
'		2=SHIPPING ADDRESS 2; ETC.


rs_Store.open sql_select_cust,conn_store,1,1
if rs_store.eof and rs_store.bof then
   rs_store.close
   response.redirect "admin_error.asp?Message_Id=101&Message_Add="&server.urlencode("The selected customer could not be found.")
end if
rs_Store.MoveFirst 
	User_id= Rs_store("User_id")
	Tax_ID= Rs_store("Tax_ID")
	Password= Rs_store("Password")
	Budget_left = Rs_store("Budget_left")
	Budget_Orig = Rs_store("Budget_Orig")
	Reward_Total = Rs_store("Reward_Total")
	Reward_Left = Rs_store("Reward_Left")
	spam = Rs_store("spam")
	CCID = Rs_store("CCID")
	CID = Rs_store("CID")
	if spam = -1 then 
		checked_spam = "checked" 
	end if
	Tax_Exempt = rs_store("Tax_Exempt")
	if Tax_Exempt = -1 then
		checked_Tax_Exempt = "checked" 
	end if
	Tax_Id=rs_store("tax_id")
	Protected_page_access = rs_store("Protected_page_access")
	if Protected_page_access then
     checked_protected = "checked"
	end if
rs_Store.Close

'CHECK	WHICH GROUPS CUSTOMER BELONGS TO

Set rs_store1 = Server.CreateObject("ADODB.Recordset")
sql_groups = "select Group_Id, Group_Name from Store_Customers_Groups where Store_id = "&Store_id&""
rs_Store1.open sql_groups,conn_store,1,1
if rs_Store1.EOF = False or rs_Store1.BOF = False then
	do while not rs_Store1.eof
		if Is_Cid_In_Group(cid,Rs_store1("Group_Id")) = true then
		 Groups = Groups&Rs_store1("Group_name")&","
		end if
		rs_store1.movenext
	loop
end if


if Groups = "" then 
	GroupsId = "None"
	Groups = "None"
end if
rs_Store1.Close 

sFormAction = "update_records_action.asp?cid="&cid
sName = "Modify_Customer"
sTitle = "Edit Customer - "&User_id
sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > Edit - "&User_id
thisRedirect = "modify_customer.asp"
sMenu = "customers"
createHead thisRedirect

Start_date =  DateAdd("m", -24, Now())
End_date =	Now()


' Get Added Groups Id
sqlCustGroups = " SELECT GROUP_ID,GROUP_NAME,GROUP_CID FROM Store_Customers_Groups WHERE "&_
						" (GROUP_CID LIKE '%,"&CCid&",%' OR GROUP_CID LIKE '"&CCid&"' OR "&_
						" GROUP_CID LIKE '%,"&CCid&"' or GROUP_CID LIKE '"&CCid&",%') AND "&_
						" STORE_ID = "&Store_Id 
set rsGroups = conn_store.execute(sqlCustGroups)
addedGroupIds =""
addedGroupNames = "None"
if not rsGroups.eof then
	while not rsGroups.eof
		if addedGroupIds <> "" then
		addedGroupIds = addedGroupIds&","&trim(rsGroups("group_id"))
		addedGroupNames = addedGroupNames&","&trim(rsGroups("group_name"))
		else
		addedGroupIds = addedGroupIds&trim(rsGroups("group_id"))
		addedGroupNames = trim(rsGroups("group_name"))
		end if
	rsGroups.movenext
	wend
end if

rsGroups.close


sql_select = "SELECT Group_ID, Group_Name from store_customers_groups where store_id = "&store_id

set groupfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,groupdata,groupfields,noRecordsgroup)

%>


<script language="JavaScript">

	function checkAddress()
	{
		selVal = "";
		for (i=0;i<document.forms['adrsel'].addrs.length;i++)
			if (document.forms['adrsel'].addrs[i].selected)
				selVal=document.forms['adrsel'].addrs[i].value;
		document.location = "Modify_customer.asp?cid=<%= cid %>&ssadr="+selVal;}

</script>
	<tr><td>
	<table width='100%' border='0' cellspacing='1' cellpadding=2>
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="20">
    </form><FORM name="Orders" method="POST" action="orders.asp?view_cid=<%= cCID %>">
		
		<input class="buttons" name="SearchOrders" type="submit" value="View Orders">
		</FORM><form method=post action=<%= sFormAction %>>
 	 </td></tr>

	 <tr>
		<td width="100%" colspan="3">
			<table width='100%' border='0' cellspacing='1' cellpadding=2>
				
				<input type="Hidden" name="Record_Type" value="0">
				
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Username</td>
			<td class="inputvalue">
						<input name="User_id" size="40" value="<%= User_id %>">
						<input name="User_id_C" value="Re|String|||" type="hidden">
			<% small_help "Username" %></td>
		</tr> 
		 
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Password</td>
			<td class="inputvalue">
						<input type="Password" name="Password" value="<%= Password %>" size="40" >
						<input name="Password_C" value="Re|String|||" type="hidden">
			<% small_help "Password" %></td>
		</tr>
		 
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Password Confirm</td>
			<td class="inputvalue">
						<input	type="Password" name="Password_Confirm" value="<%= Password %>" size="40" >
						<input name="Password_Confirm_C" value="Re|String|||" type="hidden">
			<% small_help "Password Confirm" %></td>
		</tr>
			 
		<tr bgcolor='#FFFFFF'>
			<td class="inputname">Budget</td>
			<td class="inputvalue"><%= Store_currency %>
				<input name="Budget_orig" value="<%= Budget_orig %>" size="10">
				<input name="Budget_orig_C" value="Re|Integer|||" type="hidden">
			<a class=link href=budget_view.asp?op=viewBudget&Id=<%= cid %>>View Budget Details</a><% small_help "Budget" %></td>
		  </tr>
		  
		<tr bgcolor='#FFFFFF'>
			<td class="inputname">Budget Left</td>
			<td class="inputvalue"><%= Store_currency %>
						<input name="Budget_left" value="<%= Budget_left %>" size="10">
						<input name="Budget_left_C" value="Re|Integer|||" type="hidden">
			<% small_help "Budget Left" %></td>
		</tr>
		<% if Enable_Rewards then %>
		<tr bgcolor='#FFFFFF'>
			<td class="inputname">Reward Total</td>
			<td class="inputvalue"><%= Store_currency %>
				<input name="Reward_Total" value="<%= Reward_Total %>" size="10">
				<input name="Reward_Total_C" value="Re|Integer|||" type="hidden">
			<% small_help "Reward_Total" %></td>
		  </tr>

		<tr bgcolor='#FFFFFF'>
			<td class="inputname">Reward Left</td>
			<td class="inputvalue"><%= Store_currency %>
						<input name="Reward_Left" value="<%= Reward_left %>" size="10">
						<input name="Reward_left_C" value="Re|Integer|||" type="hidden">
			<% small_help "Reward_Left" %></td>
		</tr>

		<% end if %>
			
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Promo Emails </td>
			<td class="inputvalue">
						<input class="image" type="checkbox" <%= checked_spam %> name="Spam" value="-1" >
			<% small_help "Promo Emails" %></td>
		</tr>
			
			
			
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Tax Exempt </td>
			<td class="inputvalue">
						<input class="image" type="checkbox" <%= checked_Tax_Exempt %> name="Tax_Exempt" value="-1" >
			<% small_help "Tax Exempt" %></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
			<td class="inputname">Tax ID</td>
			<td class="inputvalue">
						<input name="Tax_ID" size="40" value="<%= Tax_ID %>">
						<input name="Tax_ID_C" value="Op|String|0|20|" type="hidden">
			<% small_help "Tax_ID" %></td>
		</tr>
			<tr bgcolor='#FFFFFF'>
			<td class="inputname">Protected Access</td>
			<td class="inputvalue">
						<input class="image" type="checkbox" <%= checked_protected %> name="protected_access" value="-1" >
			<% small_help "Protected Access" %></td>
		</tr>

				
			<tr bgcolor='#FFFFFF'>
			<td class="inputname"><br><b>In Groups</b></td>
			<td class="inputvalue" colspan=2>&nbsp;</td>
			</tr>
			
					
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Matching Groups</td>
			<td class="inputvalue"><%= Groups %>
			<% small_help "Matching Groups" %></td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
			<td class="inputname" valign="top">Plus Groups </td>
			<td class="inputvalue"><%= addedGroupNames %>
			<% if noRecordsgroup = 0 then %>
			<br>
			<a href="javascript:void()" onClick="openGroupPicker('<%=inGroupsId%>')">Select Groups</a><br><textarea rows="3" name="Group_id" cols="35"><%=addedGroupIds%></textarea>
			<% end if %>
			<% small_help "Plus Groups" %></td>
			</tr>
			
			
				<tr bgcolor='#FFFFFF'>
					<td class="inputname">Customer ID</td>
			<td class="inputvalue"><%= CCid %>
			<input type="hidden" name="CCid" value="<%=CCid%>">
			<% small_help "Customer ID" %></td>
				</tr>

				<tr align=center bgcolor='#FFFFFF'>
					<td colspan=3>
						<input type="submit" class="Buttons" value="Save Customer" name="Modify_my_Account">
						&nbsp;&nbsp;&nbsp;<input OnClick='JavaScript:self.location="my_customer_base.asp"' name="Cancel" type="button" class="Buttons" value="Cancel"></td>
		</tr>
				</form>


			<table border="0" height="100%" width="100%">

			<tr bgcolor='#FFFFFF'>
					<td></td><td width="268">Address Book
									<% sql_select_addrs = "SELECT * FROM Store_Customers WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type <>	0 and store_id="&Store_Id %>
									<% rs_Store.open sql_select_addrs,conn_store,1,1 %>
									<% rs_Store.MoveFirst %> 
									<form name="adrsel">
									<% maxrt = -1 %>
									<% tota = 0 %>
									<select name="addrs" onChange="JavaScript:checkAddress();">
										<% do while not rs_Store.eof %>
											<% tota = tota + 1 %>
											<option value="<%= rs_store("Record_Type") %>"
												<% If cdbl(rs_store("Record_Type"))=cdbl(request.querystring("ssadr")) then %>
													selected
												<% End If %>
											><%= rs_store("First_name") %>&nbsp;<%= rs_store("Last_name") %>&nbsp;-&nbsp;<%= rs_store("Address1") %></option>
											<% if rs_store("Record_Type")>maxrt then %>
												<% maxrt = rs_store("Record_Type") %>
											<% End If %>
											<% rs_Store.movenext %>
										<% loop %>	
										<option value="<%= maxrt+1 %>"
											<% If cdbl(maxrt+1)=cdbl(request.querystring("ssadr")) then %>
												selected
											<% End If %>
											>Add New</option>
									</select>
									</form>
									<% rs_Store.Close %>
								</td>
				</tr>
			<tr bgcolor='#FFFFFF'>
			<td height="100%" width="50%" valign=top>
				<table border="0" width="100%">
					
							<tr bgcolor='#FFFFFF'>
					<td colSpan="2"><b>Billing Info</b></td>
					</tr>
					
							<tr bgcolor='#FFFFFF'>
					<td colSpan="2"&nbsp;></td>
					</tr>
							
							<% Record_Type = 0 %>
							<!--#include file="Display_Cust_Form.asp"-->
							
				</table>
			</td>
			
					<td height="100%" width="50%" valign=top>
				<table border="0" width="100%">
					
							<tr bgcolor='#FFFFFF'>
					<td colSpan="2" width="268"><b>Shipping Info</b></td>
					</tr>
				  

				  
					
					
							<% if tota>1 then %>
								<% if not cdbl(maxrt+1)=cdbl(request.querystring("ssadr")) then %>
									<% allow_delete = true %>
								<% else %>
									<% allow_delete = false %>
								<% end if %>
							<% else %>
								<% allow_delete = false %>
							<% end if %>
							<% if request.querystring("ssadr")<>"" then %>
								<% Record_Type = request.querystring("ssadr") %>
							<% else %>
								<% Record_Type = 1 %>
							<% end if %>
							<!--#include file="display_cust_form_Shipping.asp"-->
				
						</table>
			</td>
			</tr>
		</table>
	</td>
	</tr>
	 </table>
	</td>
	</tr>
<script language="javascript">
function openGroupPicker(inGroupsId)
{
	window.open("group_picker.asp?groupsId="+inGroupsId,"SelectGroups", "toolbar=no, location=no, width=400, height=500, status=no, scrollbars=yes, resizable=yes");
}
</script>

<% createFoot thisRedirect, 0%>
