<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
calendar=1
End_Date = FormatDateTime(now(),2)
Start_Date = dateadd("m",-1,End_Date)
sQuestion_Path = "reports/customers.htm"

if request.form="" and request.querystring("Form") <> "" then
   Form_Array = split(request.querystring("Form"),"^")
   Start_Date = Form_Array(0)
   End_Date = Form_Array(1)
end if

if request.form<>"" then
  if request.form("Start_Date")<>"" then
  	if isdate(request.form("Start_Date")) then
  		Start_Date = request.form("Start_Date")
  	end if
  end if
  if request.form("End_Date")<>"" then
  	if isdate(request.form("End_Date")) then
  		End_Date = request.form("End_Date")
  	end if
  end if
end if

End_date = DateAdd("d", 1, End_date)



Orders_type = Request.Form("Orders_type")
if Orders_type = "" then
	Orders_type = "2"
end if
	
					   
if Orders_type = "2" then
	sql_where_add = ""	
elseif Orders_type = "1" then
	sql_where_add = " and verified <>0 "
elseif Orders_type = "0" then
	sql_where_add = " and verified = 0 "
end if

if Store_Database="Ms_Sql" then
	sql_best_sellers = Replace(sql_best_sellers, "#", "'")
end if

sql_where_end = sql_where_add & " GROUP BY Item_Sku,Item_Id,Item_Name,Quantity_in_stock,verified "	


sSubmitName="SearchCustomers"
End_Date1 = DateAdd("d",-1,End_Date)
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "bestsellers"
myStructure("TableWhere") = sql_where &" Sys_created<='"&End_date&"' and Sys_created>='"&Start_Date&"'"& sql_where_end 
myStructure("ColumnList") = "item_id,item_sku,item_name,sum(quantity) as quantity_sold,quantity_in_stock"
myStructure("HeaderList") = "item_id,item_sku,item_name,quantity_sold,quantity_in_stock"
myStructure("SortList") = "item_id,item_sku,item_name,quantity_sold,quantity_in_stock"

myStructure("DefaultSort") = "quantity_sold"
myStructure("PrimaryKey") = "item_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 0
myStructure("ViewItem") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "inventory"
myStructure("FileName") = "reports_best_sellers.asp"
myStructure("FormAction") = "reports_best_sellers.asp"
myStructure("Title") = "Inventory Best Sellers"
myStructure("FullTitle") = "Inventory > Best Sellers"
myStructure("CommonName") = "Best Sellers"
myStructure("viewItem") = "item_edit.asp"

myStructure("Heading:item_id") = "PK"
myStructure("Heading:item_sku") = "SKU"
myStructure("Heading:item_name") = "Name"
myStructure("Heading:quantity_sold") = "Sold"
myStructure("Heading:quantity_in_stock") = "Quantity"

myStructure("Format:item_id") = "STRING"
myStructure("Format:item_sku") = "STRING"
myStructure("Format:item_name") = "STRING"
myStructure("Format:quantity_sold") = "INT"
myStructure("Format:quantity_in_stock") = "INT"
myStructure("Form") = Start_Date&"^"&End_Date

%>
<script language="JavaScript">
	function showhideadvanced(){
	if (document.forms[0].AdvancedSearch.checked)
  	{
  	document.all.DS.style.display = 'block';
  	}
  	else
  	{
  	document.all.DS.style.display = 'none';
  	}
  }
</script>
<!--#include file="head_view.asp"-->
<input type=hidden name=records value=<%= RowsPerPage %>>

<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='13'><table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
	
	<tr bgcolor='#FFFFFF'><td class="inputname" width='40%'>Status</td>
		<td class="inputvalue">
			<select name="Orders_type">
				<option value="2"
				<% if Orders_type = "2" then %>
					selected
				<% End If %>
				>All</option>
				<option value="1"
				<% if Orders_type = "1" then %>
					selected
				<% End If %>
				>Verified</option>
				<option value="0"
				<% if Orders_type = "0" then %>
					selected
				<% End If %>
				>Non Verified</option>
			</select>
			<% small_help "Status" %>
		</td>
	</tr>
	
	
	<tr bgcolor='#FFFFFF'>
		<td class="inputname">Between</td>
		<td class="inputvalue">
    <SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
			<input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
			<A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			and <input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date1,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
			<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>

<% small_help "Between" %></td></tr>

	<tr bgcolor='#FFFFFF'><td colspan=3 align=center><input name="<%= sSubmitName %>" type="submit" class="Buttons" value="Search Best Sellers"></td></tr>
	
</table></TD></TR>

<!--#include file="list_view.asp"-->

<%
createFoot thisRedirect, 2
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");

</script>

