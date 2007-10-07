<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
sQuestion_Path = "advanced/insurance.htm"
sTextHelp="insurance/percentage_insurance.doc"

op=Request.QueryString("op")
if op<>"" then
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlSelShipping="select * from Store_Insurance_class_all where " & _
						"Insurance_class=5 AND Store_id="&Store_id & " and " & _
						"insurance_method_Id=" & Request.QueryString("Id")
	rs_store.open sqlSelShipping,conn_store,1,1
	InsuranceMethodName=rs_store("Insurance_Method_Name")
	MatrixLow=formatnumber(rs_store("Matrix_Low"),2)
	MatrixHigh=formatnumber(rs_store("Matrix_High"),2)
	BaseFee=formatnumber(rs_store("Base_Fee"),2)
	rs_store.close
end if

sMenu="shipping"
sFormAction = "update_records_action.asp"
sName = "Add_Class_5"
sCommonName="Shipping Insurance"
sCancel="insurance_class5_list.asp"

if request.querystring("Id")="" then
        sFullTitle = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > <a href=insurance_class5_list.asp class=white>% Total Order</a> > Add"
        sTitle = "Add % Total Order Shipping Insurance"
else
        sFullTitle = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > <a href=insurance_class5_list.asp class=white>% Total Order</a> > Edit - "&InsuranceMethodName
        sTitle = "Edit % Total Order Shipping Insurance - "&InsuranceMethodName
end if
thisRedirect = "insurance_class5_add.asp"
createHead thisRedirect
%>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Insurance_Method_Id" value="<%=Request.QueryString("Id")%>">
<input type="hidden" name="Insurance_Form_Back_To" value="insurance_class5_list.asp">
<input type="hidden" name="Add_What_Insurance_Class" value="5">





				<TR bgcolor='#FFFFFF'>
					<td width="20%"class="inputname"><B>Name</b></td>
				<td width="80%" class="inputvalue">
						<input type="text" name="Insurance_Method_Name" size="60" maxlength=200 value="<%= InsuranceMethodName %>">*
						<input type="hidden" name="Insurance_Method_Name_C" value="Re|String|0|200|||Insurance Method">
						<% small_help "Name" %></td>
			</tr>
	  
				<TR bgcolor='#FFFFFF'>
				<td width="20%"class="inputname"><B>Fee</b></td>
				<td width="80%" class="inputvalue">
						<input type="text" name="Base_Fee" size="10" value="<%= BaseFee %>" onKeyPress="return goodchars(event,'0123456789.')">*
						<input type="hidden" name="Base_Fee_C" value="Re|Integer|||||Fee">%
						<% small_help "Fee" %></td>
			</tr>
			<TR bgcolor='#FFFFFF'>
				<td width="20%"class="inputname"><B>Total Between</b></td>
				<td width="80%" class="inputvalue">
					<%= Store_Currency %><input type="text" name="Matrix_Low" size="10" value="<%= MatrixLow %>" onKeyPress="return goodchars(event,'0123456789.')">*
						<input type="hidden" name="Matrix_Low_C" value="Re|Integer|||||Matrix Low">
						and
						<%= Store_Currency %><input type="text" name="Matrix_High" size="10" value="<%= MatrixHigh %>" onKeyPress="return goodchars(event,'0123456789.')">*
						<input type="hidden" name="Matrix_High_C" value="Re|Integer|||||Matrix High">
						<% small_help "Total Between" %></td>
			</tr>

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Insurance_Method_Name","req","Please enter a name for this shipping method.");
 frmvalidator.addValidation("Base_Fee","req","Please enter a base fee.");
 frmvalidator.addValidation("Matrix_Low","req","Please enter a low range.");
 frmvalidator.addValidation("Matrix_High","req","Please enter a high range.");
 </script>
