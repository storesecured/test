<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->

<%

op=Request.QueryString("op")
sFullTitle = "General > <a href=tax_list.asp class=white>Taxes</a> > "
if op="edit" then
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlTaxes="select * from store_Zips where store_id=" & store_id & " and Zip_Range_ID=" & Request.QueryString("Zip_Range_ID")
	rs_Store.open sqlTaxes,conn_store,1,1
		Zip_Start=rs_store("Zip_Start")
		Zip_End=rs_store("Zip_End")
		DepartmentIds=rs_store("Department_Ids")
		TaxRate=rs_store("TaxRate")
		Name=rs_store("Zip_Name")
		Tax_Shipping=rs_Store("Tax_Shipping")
	rs_store.close
	sFullTitle=sFullTitle&"Edit Zipcode Tax - "&Name
        sTitle = "Edit Zipcode Tax - "&Name
else
	sFullTitle=sFullTitle&"Add Zipcode Tax"
        sTitle = "Add Zipcode Tax"
end if

sInstructions="All ship to zip codes between zip start and zip end will be charged the tax rate below for each item purchased from the selected departments."
sTextHelp="taxes/general-tax-zip.doc"

sFormAction = "tax_action.asp"
sName = "tax"
sCommonName="Zipcode Tax"
sCancel="tax_list.asp"
thisRedirect = "tax_add.asp"
sFormName = "Add_Tax"
sSubmitName = "Add_Tax"
sMenu="general"
createHead thisRedirect
%>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Zip_Range_ID" value="<%=Request.QueryString("Zip_Range_ID")%>">


<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Name</b></td>
				<td width="90%" class="inputvalue">
					<input maxLength="20" name="Name" size="60" maxlength=50
							<% if op="edit" then %>
								value="<%= Name %>"
						<% else %>
								value=""
						<% end if %>
						>
								 <input type="hidden" name="Name_C" value="Re|String|0|50|||Name">
						<% small_help "Name" %></td>
					</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Zip Between</b></td>
				<td width="90%" class="inputvalue">
					<input maxLength="5" name="Zip_Start" size="5" onKeyPress="return goodchars(event,'0123456789')"
							<% if op="edit" then %>
								value="<%= Zipcode_Format_Function(Zip_Start) %>"
						<% else %>
								value=""
						<% end if %>
						>
										<input type="hidden" name="Zip_Start_C" value="Re|Integer|0|99999|||Zip Start">
						and
						<input maxLength="5" name="Zip_End" size="5" onKeyPress="return goodchars(event,'0123456789')"
							<% if op="edit" then %>
								value="<%= Zipcode_Format_Function(Zip_End) %>"
						<% else %>
								value=""
						<% end if %>
						>
										<input type="hidden" name="Zip_End_C" value="Re|Integer|0|99999|||Zip End">
						<% small_help "Zip Between" %></td>
					</tr>
					
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Tax Rate</b></td>
				<td width="90%" class="inputvalue">
					<input maxLength="8" name="TaxRate" size="4"  onKeyPress="return goodchars(event,'0123456789.')"
							<% if op="edit" then %>
								value="<%= formatnumber(TaxRate,4) %>"
						<% else %>
								value="<%= formatnumber(0,4) %>"
						<% end if %>
						><input type="hidden" name="TaxRate_C" value="Re|Integer|0|100|||Tax Rate">%
						<% small_help "Tax Rate" %></td>
					</tr>
		
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Departments</b>
						</td>
				<td width="90%" class="inputvalue">
						<%= create_dept_list ("Department_IDs",DepartmentIds,4,"-1|{}|All|{new}|") %>
						
					<input type="hidden" name="Department_IDs_C" value="Re|String|||||Departments">
					<% small_help "Departments" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Tax Shipping</b></td>
				<td width="90%" class="inputvalue">
							<% if Tax_Shipping then
								checked = "checked"
							else
								checked = ""
							end if %>
							<input class="image" type="checkbox" <%= checked %> name="Tax_Shipping">

						<% small_help "Tax Shipping" %></td>
					</tr>


<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Name","req","Please enter a name for this tax rate.");
 frmvalidator.addValidation("TaxRate","req","Please enter a tax rate.");
 frmvalidator.addValidation("Department_IDs","req","Please select the departments to be charged this tax rate.");
 frmvalidator.addValidation("Zip_Start","req","Please enter the zip code range to start at for this tax rate will apply to.");
 frmvalidator.addValidation("Zip_End","req","Please enter the zip code range to end at for this tax rate will apply to.");
</script>
