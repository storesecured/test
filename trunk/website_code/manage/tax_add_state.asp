<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->

<%

op=Request.QueryString("op")
sCommonName="State Tax"
sCancel="tax_list.asp"
sFullTitle = "General > <a href=tax_list.asp class=white>Taxes</a> > "
if op="edit" then
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlTaxes="select * from store_State_Tax where store_id=" & store_id & " and State_Tax_ID=" & Request.QueryString("State_Tax_ID")
	rs_Store.open sqlTaxes,conn_store,1,1
		State=rs_store("State")
		DepartmentIds=rs_store("Department_Ids")
		TaxRate=rs_store("TaxRate")
		Tax_Shipping=rs_Store("Tax_Shipping")
	rs_store.close
	sFullTitle=sFullTitle&"Edit State Tax - "&State
        sTitle = "Edit State Tax - "&State
else
	sFullTitle=sFullTitle&"Add State Tax"
        sTitle = "Add State Tax"
end if

sInstructions="All selected ship to states will be charged the tax rate below for each item purchased from the selected departments."    
sTextHelp="taxes/general-tax-state.doc"

sFormAction = "tax_action.asp"
sName = "tax"
thisRedirect = "tax_add_state.asp"
sFormName = "Add_Tax_State"
sSubmitName = "Add_Tax_State"
sMenu="general"
createHead thisRedirect
%>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="State_Tax_ID" value="<%=Request.QueryString("State_Tax_ID")%>">

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
				<td width="10%" class="inputname"><b>State / Province</b></td>
				<td width="90%" class="inputvalue">
					<% sql_select = "SELECT State,State_Long from Sys_States order by State"
						set myfields=server.createobject("scripting.dictionary")
						Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
						selected_flag="false"
						response.write "<select multiple size=4 name=State>"
						response.write "<option value=''"
            if State="" then
               response.write " selected"
            end if
            response.write ">Select valid states for this tax rate</option>"
						if noRecords = 0 then
							 FOR rowcounter= 0 TO myfields("rowcount")
								if op="edit" then
									StateArray=split(State,",")
									Found=false
									if Is_In_Collection(State,mydata(myfields("state"),rowcounter),",") then
										Found=true
									end if
									if Found then
										selected = "selected"
									else
										selected=""
									end if
								else
									selected = ""

								end if
								response.write "<option "&selected&" value='"&mydata(myfields("state"),rowcounter)&"'>"&mydata(myfields("state_long"),rowcounter)&"</option>"
							Next
						 End If %>
					</select>
			 <input type="hidden" name="State_C" value="Re|String|0|50|||State">
						<% small_help "State" %></td>
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
 frmvalidator.addValidation("State","req","Please select the states that this tax rate will apply to.");
</script>
