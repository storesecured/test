<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->
<!--#include file="include/country_list.asp"-->
<%

op=Request.QueryString("op")
sCommonName="Country Tax"
sCancel="tax_list.asp"
sFullTitle = "General > <a href=tax_list.asp class=white>Taxes</a> > "
if op="edit" then
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlTaxes="select * from store_Country_Tax where store_id=" & store_id & " and Country_Tax_ID=" & Request.QueryString("Country_Tax_ID")
	rs_Store.open sqlTaxes,conn_store,1,1
		Country=rs_store("Country")
		DepartmentIds=rs_store("Department_Ids")
		TaxRate=rs_store("TaxRate")
		Tax_Shipping=rs_Store("Tax_Shipping")
	rs_store.close
	sFullTitle=sFullTitle&"Edit Country Tax - "&Country
        sTitle = "Edit Country Tax - "&Country
else
	sFullTitle=sFullTitle&"Edit Country Tax"
        sTitle = "Add Country Tax"
end if

sInstructions="All selected ship to countries will be charged the tax rate below for each item purchased from the selected departments."
sTextHelp="taxes/general-tax-country.doc"

sFormAction = "tax_action.asp"
sName = "tax"
thisRedirect = "tax_add_country.asp"
sFormName = "Add_Tax_Country"
sSubmitName = "Add_Tax_Country"
sMenu="general"
createHead thisRedirect
%>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Country_Tax_ID" value="<%=Request.QueryString("Country_Tax_ID")%>">

				<tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Tax Rate</b></td>
				<td width="90%" class="inputvalue">
					<input maxLength="8" name="TaxRate" size="4" onKeyPress="return goodchars(event,'0123456789.')"
							<% if op="edit" then %>
								value="<%= formatnumber(TaxRate,4) %>"
						<% else %>
								value="<%= formatnumber(0,4) %>"
						<% end if %>
						><input type="hidden" name="TaxRate_C" value="Re|Integer|0|100|||Tax Rate">%
						<% small_help "Tax Rate" %></td>
					</tr>
          <tr bgcolor='#FFFFFF'>
				<td width="10%" class="inputname"><b>Countries</b></td>
				<td width="90%" class="inputvalue">
					<%= create_country_list ("Country",Country,4,"All Countries:") %>
					
			 <input type="hidden" name="Country_C" value="Re|String|0|50|||Country">
						<% small_help "Country" %></td>
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
 frmvalidator.addValidation("TaxRate","req","Please enter a tax rate.");
 frmvalidator.addValidation("Department_IDs","req","Please select the departments to be charged this tax rate.");
 frmvalidator.addValidation("Country","req","Please select the countries that this tax rate will apply to.");
</script>
