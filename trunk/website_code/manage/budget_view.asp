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

End_Date1=DateAdd("d",1,End_Date)
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Budget_History"
myStructure("TableWhere") = "cid="&request.querystring("Id")&" and budget_date<='"&End_Date1&"' and budget_date>='"&Start_Date&"'"
myStructure("ColumnList") = "budgetid,budget_date,budget_amount,notes"
myStructure("DefaultSort") = "budget_date"
myStructure("PrimaryKey") = "budgetid"
myStructure("Level") = 5
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = "my_customer_base.asp"
myStructure("Menu") = "customers"
myStructure("FileName") = "budget_view.asp?Id="&request.querystring("Id")
myStructure("FormAction") = "budget_view.asp?Id="&request.querystring("Id")
myStructure("Title") = "Customer Budget History"
myStructure("CommonName") = "budget history"
myStructure("NewRecord") = ""
myStructure("Heading:budgetid") = "PK"
myStructure("Heading:budget_date") = "Date"
myStructure("Heading:budget_amount") = "Amount"
myStructure("Heading:notes") = "Notes"
myStructure("Format:budget_date") = "STRING"
myStructure("Format:budget_amount") = "STRING"
myStructure("Format:notes") = "STRING"
myStructure("Form") = Start_Date &"^"&End_Date
%>
<!--#include file="head_view.asp"-->
<tr><td colspan=10><table width='100%' border='0' cellspacing='1' cellpadding=2><tr bgcolor='#FFFFFF'>
		<td class="inputname"><b>Between</b></td>
		<td class="inputvalue"><SCRIPT LANGUAGE="JavaScript" ID="jscal1">
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
	<tr bgcolor='#FFFFFF'><td colspan=3 align=center><input type="submit" value="Search"></td></tr>
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
