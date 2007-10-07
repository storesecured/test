<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


op=Request.QueryString("op")
if op<>"" then
	' IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlSelShipping="select * from Store_Insurance_class_all where " & _
						"insurance_class=1 AND Store_id="&Store_id & " and " & _
						"insurance_method_Id=" & Request.QueryString("Id")
	rs_store.open sqlSelShipping,conn_store,1,1
	InsuranceMethodName=rs_store("Insurance_Method_Name")
	BaseFee=formatnumber(rs_store("Base_Fee"),2)
	rs_store.close
end if

sTextHelp="insurance/flat_insurance.doc"

sFormAction = "update_records_action.asp"
sName = "Add_Class_1"
sCancel="insurance_class1_list.asp"
sCommonName="Shipping Insurance"
if request.querystring("Id")="" then
        sTitle = "Add Flat Fee Shipping Insurance"
        sFullTitle = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > <a href=insurance_class1_list.asp class=white>Flat Fee</a> > Add"
else
        sTitle = "Edit Flat Fee Shipping Insurance - "&InsuranceMethodName
        sFullTitle = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > Flat Fee<a href=insurance_class1_list.asp class=white>Flat Fee</a> > Edit - "&InsuranceMethodName
end if
thisRedirect = "insurance_class1_add.asp"
sMenu = "shipping"
sQuestion_Path = "advanced/insurance.htm"
createHead thisRedirect


%>


<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Insurance_Method_Id" value="<%=Request.QueryString("Id")%>">
<input type="hidden" name="Insurance_Form_Back_To" value="insurance_class1_list.asp">
<input type="hidden" name="Add_What_Insurance_Class" value="1">


	

<TR bgcolor='#FFFFFF'>
					<td width="20%"class="inputname"><B>Name</b></td>
					<td width="80%" class="inputvalue">
							<input type="text" name="Insurance_Method_Name" size="60" maxlength=200 value="<%= InsuranceMethodName %>" maxlength=200>*
						<input type="hidden" name="Insurance_Method_Name_C" value="Re|String|0|200|||Shipping Method">
						<% small_help "Name" %></td>
					</tr>
			 
				<TR bgcolor='#FFFFFF'>
					<td width="20%"class="inputname"><B>Flat Fee</b></td>
					<td width="80%" class="inputvalue">
							<%= Store_Currency %><input type="text" name="Base_Fee" size="10" value="<%= BaseFee %>" onKeyPress="return goodchars(event,'0123456789.')">*
						<input type="hidden" name="Base_Fee_C" value="Re|Integer|||||Flat Fee">
						<% small_help "Flat Fee" %></td>
					</tr>




<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Insurance_Method_Name","req","Please enter a name for this insurance method.");
 frmvalidator.addValidation("Base_Fee","req","Please enter a flat insurance amount.");

</script>

