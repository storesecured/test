<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

Root_Position = InStr(Request.ServerVariables("URL"), "/Admin")
Store_Root = left(Request.ServerVariables("URL"),Root_Position)


sFormAction = "Orders.asp"
sName = "Orders"
sTitle = "Home"
sSubmitName = "Store_Activation_Update"
thisRedirect = "admin_home.asp"
sTopic = "Home"
addPicker=1
sMenu="home"
createHead thisRedirect

Start_date =  DateAdd("m", -12, Now())
End_date =	Now()

sql_select = "wsp_todo_list "&Store_Id

rs_store.open sql_select,conn_store,1,1
do while not rs_store.eof
	this_column=rs_store("this_column")
	this_value=rs_store("this_value")
	select case this_column
		case "dept"
			 count_dept=this_value
		case "items"
			 count_items=this_value
		case "shipping"
			 count_shipping=this_value
		case "zips"
			 count_zips=this_value
		case "state"
			 count_state=this_value
		case "country"
			 count_country=this_value
		case "template"
			 count_template=this_value
		case "pages"
			 count_pages=this_value
		case "page_content"
			 page_content=this_value
		case "verified"
			 notshipped=this_value
		case "unverified"
			 unveriforders=this_value
		case "support"
			 support_id=this_value
		case "customers"
			 customers=this_value
	end select
	rs_store.movenext
loop
rs_store.close

 %>
 
	

<tr valign=top><td>
<% if No_Ecommerce = 0 then %>
<table cellpadding=5 cellspacing=2 width="100%">

  <tr align='center' bgcolor='#0069B2' class='white'>
    <td width=18%><font class=white><B>Orders</B></font></td>
    <td width=18%><font class=white><B>Inventory</B></font>
    
    
  </td></tr>
  <tr bgcolor='#FFFFFF'>
    <td valign=top align=center>
		
		<FORM name="Orders" method="POST" action="orders.asp">
		<input type=hidden name="Show" value=1>
		<input name="Start_Date" type=hidden value="<%= FormatDateTime(Start_date,2) %>">
		<input name="End_Date" type=hidden value="<%= FormatDateTime(End_date,2) %>">
		<input name="ShippingStatus" type=hidden value=2>
		<input class="buttons" name="SearchOrders" type="submit" value="<%=notshipped %> Unshipped Order(s)">
		</FORM>
		

		
		<FORM name="Orders" method="POST" action="orders.asp">
    <input type=hidden name="Show" value=2>
		<input name="Start_Date" type=hidden value="<%= FormatDateTime(Start_date,2) %>">
		<input name="End_Date" type=hidden value="<%= FormatDateTime(End_date,2) %>">
		<input name="ShippingStatus" type=hidden value=0>
		<input class="buttons" name="SearchOrders" type="submit" value="<%= unveriforders %> Unverified Order(s)">
		</form>


		
    </td>
	<td valign=top  align=center>
	<%
        if count_items=0 then
	       item_location="item_basic_edit.asp"
	elseif count_items<=250 then
	       item_location="edit_items.asp?Form="
	else
	       item_location="edit_items.asp"
	end if
	if count_dept=0 then
	       dept_location="store_dept_basic.asp"
	else
	       dept_location="department_manager.asp"
	end if
	%>
        <input type="button" class="Buttons" value="<%= count_items %> Item(s)" name="Add" OnClick='JavaScript:self.location="<%= item_location %>"'><br><br>
	<input type="button" class="Buttons" value="<%= count_dept %> Department(s)" name="Add" OnClick='JavaScript:self.location="<%= dept_location %>"'><br><br>
	
	</td>
	
    
	

  </tr>
<tr align='center' bgcolor='#0069B2' class='white'>
    <td width=18%><font class=white><B>Customers</B></font></td>
    <td width=18%><font class=white><B>Support</B></font></td>
    
  </td></tr>
  <tr bgcolor='#FFFFFF'>
  <td align=center>
	<input type="button" class="Buttons" value="<%= customers %> Customer(s)" name="Add" OnClick='JavaScript:self.location="my_customer_base.asp"'><br><br>
	
    
	</td>
	<td valign=top align=center>
		<% if support_id<>"" then
			sSupport=1
		else
			sSupport="No"
		end if %>
			
		<input type="button" class="Buttons" value="<%= sSupport %> Question(s) from Support" name="Add" OnClick='JavaScript:self.location="support_edit.asp?Support_Id=<%= support_id %>"'><br><br>
		
		</td>
    
    
</tr>

  </table>
  <table cellpadding=5 cellspacing=2 width="100%">
  <tr bgcolor='#FFFFFF'><td><B><a href=http://videos.storesecured.com/gettingstarted/getting_started.html target=_blank>Need Help?</a></b> View the getting started tutorial for the basics on building your site. <a href=http://videos.storesecured.com/gettingstarted/getting_started.html target=_blank>Click here</a></td></tr></table>
<% else %>
Your site is available on the web at <%= Site_Name %>
<% end if %>
  </td><td width=250><table border=0 cellspacing=2 cellpadding=5 height=100%>
    <tr align='center' bgcolor='#0069B2' class='white' height='100%'><td><font class=white><B>Announcements</B></font></td></tr>
	<tr bgcolor='#FFFFFF'><td height=210 valign=top>
      <B>9/11</b> New Features added <a href=features.asp class=link>more details...</a>
      <BR><BR><B>5/23</b> PayPal Express Checkout free processing <a href=paypal_christmas.asp class=link>more details...</a>
      <BR><BR><B>2/6</b> New store owner forum <a href=forum.asp class=link>more details...</a>
      <BR><BR><B>3/14</b> New editor video tutorials <a href=http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/kb-5548564386402645416/view_question!PAGETYPE?sq=editor%2btutorial&sf=101113&sg=1&st=379136&documentid=184460&action=view class=link target=_blank>more details...</a>


  </td></tr></table></td></tr>





	  <% createFoot thisRedirect, 0%>
