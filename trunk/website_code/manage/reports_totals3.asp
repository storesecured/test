<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

'COUNT DEPARTMENTS
sql_Department = "SELECT Count(Department_ID) AS CountOfDepartment_ID FROM Store_Dept where Store_id = "&Store_id&" and Visible=1;" 
rs_Store.open sql_Department,Conn_store,1,1 
	if rs_Store.bof = false then
		rs_Store.MoveFirst	  
		' get the verified 
		Number_Of_departments = Rs_store("CountOfDepartment_ID")
	else
		Number_Of_departments = 0
	end if
rs_Store.Close

'COUNT DEPARTMENTS
sql_Department = "SELECT Count(Department_ID) AS CountOfDepartment_ID FROM Store_Dept where Store_id = "&Store_id&" and Visible=0;" 
rs_Store.open sql_Department,Conn_store,1,1 
	if rs_Store.bof = false then
		rs_Store.MoveFirst	  
		' get the verified 
		Number_Of_departments_Hidden = Rs_store("CountOfDepartment_ID")
	else
		Number_Of_departments_Hidden = 0
	end if
rs_Store.Close

'COUNT ITEMS
sql_Item = "SELECT Count(Item_ID) AS CountOfItem_ID FROM Store_Items where Store_id = "&Store_id&" and Show = 1;" 
rs_Store.open sql_Item,Conn_store,1,1 
	if rs_Store.bof = false then
		rs_Store.MoveFirst	  
		' get the verified 
		Number_Of_Active_Items = Rs_store("CountOfItem_ID")
	else
		Number_Of_Active_Items = 0
	end if
rs_Store.Close

sql_Item = "SELECT Count(Item_ID) AS CountOfItem_ID FROM Store_Items where Store_id = "&Store_id&" and Show = 0;" 
rs_Store.open sql_Item,Conn_store,1,1 
	if rs_Store.bof = false then
		rs_Store.MoveFirst	  
		' get the verified 
		Number_Of_Non_Active_Items = Rs_store("CountOfItem_ID")
	else
		Number_Of_Non_Active_Items = 0
	end if
rs_Store.Close

Number_Of_Items = Number_Of_Active_Items + Number_Of_Non_Active_Items
Number_Of_departments_Total=Number_Of_departments_Hidden+Number_Of_departments

sFormAction = "Reports_Totals.asp"
sName = "Reports_Totals"
sTitle = "Inventory Report"
sFullTitle = "Inventory > Report"
thisRedirect = "Reports_Totals.asp"
sMenu = "inventory"
sQuestion_Path = "reports/inventory_totals.htm"
createHead thisRedirect
%>


	
				
				<tr align='center' bgcolor='#0069B2' class='white'>
					<td width="273" class=tablehead><font class=white><b>Inventory&nbsp;</b></font></td>
					<td width="278" class=tablehead><font class=white><b>Total</b></font></td>
					<td width="278" class=tablehead><font class=white><b>Visible</b></font></td>
					<td width="278" class=tablehead><font class=white><b>Not Visible</b></font></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="273" class=0>Items&nbsp; </td>
					<td width="278" class=0><%= Number_Of_Items %></td>
					<td><%= Number_Of_Active_Items %></td>
					<td><%= Number_Of_Non_Active_Items %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="273" class=1>Departments&nbsp; </td>
					<td width="278" class=1><%= Number_Of_departments_Total %></td>
					<td width="278" class=1><%= Number_Of_departments %></td>
					<td width="278" class=1><%= Number_Of_departments_hidden %></td>
				</tr>
			
<% createFoot thisRedirect, 0%>
