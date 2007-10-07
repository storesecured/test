<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
calendar=1
End_Date = FormatDateTime(now(),2)
Start_Date = dateadd("m",-1,End_Date)

if request.form="" and request.querystring("Form") <> "" then
   Form_Array = split(request.querystring("Form"),"^")
   Start_Date = Form_Array(0)
   End_Date = Form_Array(1)
end if

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

sQuestion_Path = "reports/coupons_1.htm"

End_Date1 = DateAdd("d",1,End_Date)
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_purchases"
myStructure("TableWhere") = "Sys_Created<='"&End_Date1&"' and Sys_Created>='"&Start_Date&"' and coupon_id<>'' and purchase_completed=1"
myStructure("ColumnList") = "oid,oid as orderid,sys_created,verified,coupon_id,coupon_amount,shipfirstname,shiplastname,grand_total"
myStructure("HeaderList") = "oid,oid as orderid,sys_created,verified,coupon_id,coupon_amount,shipfirstname,shiplastname,grand_total"
myStructure("DefaultSort") = "oid"
myStructure("PrimaryKey") = "oid"
myStructure("Level") = 0
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 0
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "reports_coupons.asp"
myStructure("FormAction") = "reports_coupons.asp"
myStructure("Title") = "Coupon Report"
myStructure("FullTitle") = "Marketing > <a href=coupon_manager.asp class=white>Coupons</a> > Report"
myStructure("CommonName") = "Coupon"
myStructure("NewRecord") = ""
myStructure("Heading:oid") = "PK"
myStructure("Heading:oid as orderid") = "Order Id"
myStructure("Heading:verified") = "Verified"
myStructure("Heading:coupon_id") = "Coupon Id"
myStructure("Heading:coupon_amount") = "Coupon Amount"
myStructure("Heading:shipfirstname") = "First Name"
myStructure("Heading:shiplastname") = "Last Name"
myStructure("Heading:grand_total") = "Grand Total"
myStructure("Heading:sys_created") = "Date"
myStructure("Format:oid as orderid") = "STRING"
myStructure("Format:verified") = "LOOKUP"
myStructure("Format:coupon_id") = "STRING"
myStructure("Format:coupon_amount") = "CURR"
myStructure("Format:shipfirstname") = "STRING"
myStructure("Format:shiplastname") = "STRING"
myStructure("Format:grand_total") = "CURR"
myStructure("Format:sys_created") = "DATE"
myStructure("Link:coupon_id") = "new_coupon.asp?op=edit&CouponId=THISFIELD"
myStructure("Link:oid as orderid") = "order_details.asp?op=edit&Id=THISFIELD"
myStructure("Lookup:verified") = "False:No^True:Yes"



myStructure("Form") = Start_Date &"^"&End_Date
%>
<!--#include file="head_view.asp"-->

<TR><td colspan=10><table><TR bgcolor='#FFFFFF'>
		<td class="inputname" width="20%"><b>Between</b></td>
		<td class="inputvalue" width="80%"><SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
			<input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
      <A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			<input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
			<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>

		<input type=hidden value=<%= bid %> name=bid>
		<% small_help "Between" %></td>
	</tr>
	<TR bgcolor='#FFFFFF'><td colspan=3 align=center><input type="submit" value="Search Coupons"></td></tr>
</table></TD></TR>

<!--#include file="list_view.asp"-->

<%
	if Request.QueryString("Delete_Id") <> "" then

	end if 
createFoot thisRedirect, 0
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");

</script>

