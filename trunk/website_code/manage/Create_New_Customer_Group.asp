<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

' CHECK IF IN EDIT MODE, IF YES RETRIEVE CURRENT VALUES
' FROM THE DATABASE
op=Request.QueryString("op")

if op="edit" then
	sqlGroup="select * from Store_Customers_Groups where Group_Id=" & Request.QueryString("Id") & " and store_id=" & store_id
	rs_Store.open sqlGroup,conn_store,1,1
		GroupName=rs_store("Group_Name")
		GroupCountry=rs_store("Group_Country")
		GroupCity=rs_store("Group_City")
		GroupBudgetMin=rs_store("Group_Budget_Min")
		GroupCompany=rs_store("Group_Company")
		GroupPurchaseHistory=rs_store("Group_Purchase_History")
		GroupCid=rs_store("Group_Cid")
		GroupDept=rs_store("Group_Dept")
	rs_store.close
end if

if op="edit" then
	sTitle = "Edit Customer Group - "&GroupName
	sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > <a href=customers_groups.asp class=white>Groups</a> > Edit - "&GroupName
else
	 sTitle ="Add Customer Group"
	 sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > <a href=customers_groups.asp class=white>Groups</a> > Add"
end if
sCommonName="Customer Group"
sCancel="customers_groups.asp"
sFormAction = "update_records_action.asp"
thisRedirect = "create_new_customer_group.asp"
sMenu = "customers"
sQuestion_Path = "advanced/customer_groups.htm"
createHead thisRedirect

%>


<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Customer_Group_Id" value="<%=Request.QueryString("Id")%>">
<input type="hidden" name="Create_Customer_Group" value="Process">


	 
		



					<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Name</b></td>
						<td width="60%" class="inputvalue">
						<input type="text" name="Group_name" size="60" value="<%= GroupName %>" maxlength=250>
							<INPUT type="hidden"  name=Group_name_C value="Re|String|0|250|||Name">
                                   <% small_help "Name" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Company name</b></td>
						<td width="60%" class="inputvalue">
							<input type=text name="Group_Company" value="<%= GroupCompany %>" size=60>
							<INPUT type="hidden"  name=Group_Company_C value="Re|String|||||Company Name"><BR>Use All to specify all companies.
                                   <% small_help "Company Name" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Purchase Minimum</b></td>
						<td width="60%" class="inputvalue">
							<%= Store_Currency %><input type="text" name="Group_Purchase_History" onKeyPress="return goodchars(event,'0123456789.')" size="9" value="<%= GroupPurchaseHistory %>">
							<INPUT type="hidden"  name=Group_Purchase_History_C value="Re|Integer|||||Purchase Minimum">
                                   <% small_help "Purchase Minimum" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Remaining Budget</b></td>
						<td width="60%" class="inputvalue">
							<%= Store_Currency %><input type="text" name="Group_budget_min" onKeyPress="return goodchars(event,'0123456789.')" size="9" value="<%= GroupBudgetMin %>">
							<input type="hidden" value="All" name="group_city"></td><INPUT type="hidden" name=Group_budget_min_C value="Re|Integer|||||Remaining Budget">
                                   <% small_help "Remaining Budget" %></td>
				</tr>
	  
				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Country</b></td>
						<td width="60%" class="inputvalue">
							<select size="4" name="group_country" multiple>
							<option value="All" 
							<% if op<>"edit" then
								response.write "selected"
							else
								if instr(1,GroupCountry,"All")>0 then
									response.write "selected"
								end if
							end if
							response.write ">All Countries</option>"
							sql_region = "SELECT Country FROM Sys_Countries where Country <> 'All Countries' ORDER BY Country;"
							set myfields=server.createobject("scripting.dictionary")
							Call DataGetrows(conn_store,sql_region,mydata,myfields,noRecords)

								if noRecords = 0 then
								FOR rowcounter= 0 TO myfields("rowcount")

									if op="edit" then
										CountriesArray=split(GroupCountry,",") 
										Found=false 
										For i=0	to  ubound(CountriesArray) 
											CountryElem=CountriesArray(i) 
											CountryElem=trim(CountryElem) 
											Country=mydata(myfields("country"),rowcounter)
											Country=trim(Country) 
											If CountryElem = Country Then 
												Found=true 
												exit for 
											End If 
										Next 
										if Found then 
											selected = "selected" 
										else 
											selected="" 
										end if 
									end if 
									response.write "<Option "&selected&" value='"&mydata(myfields("country"),rowcounter)&"'>"&mydata(myfields("country"),rowcounter)
								Next
								end if
							response.write "</select>"
                            small_help "Country" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Dept Purchases</b></td>
						<td width="60%" class="inputvalue">
							<select size="4" name="Group_Dept" multiple>

              <option value=""
								<% if op<>"edit" then
									response.write "selected"
								else
									if (GroupDept="") or isNull(GroupDept) then
										response.write "selected"
									end if
								end if
								response.write ">All Depts</option>"

                                sql_select = "SELECT Department_ID, Department_Name,Full_Name,Belong_to,Last_Level from store_dept where Department_ID <> 0 and Store_id="&Store_id&" order by Full_Name"
                                set myfields=server.createobject("scripting.dictionary")
								Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

								if noRecords = 0 then
								FOR rowcounter= 0 TO myfields("rowcount")

									if op="edit" and not isNull(GroupDept) then
										DeptArray=split(GroupDept,",")
										Found=false 
										For i=0	to  ubound(DeptArray) 
											DeptElem=DeptArray(i) 
											DeptElem=trim(DeptElem) 
											Dept=mydata(myfields("department_id"),rowcounter)
											Dept=trim(Dept)
											If DeptElem = Dept Then
												Found=true 
												exit for 
											End If
										Next
										if Found then
											selected = "selected"
										else
											selected=""
										end if
									end if
									response.write "<Option "&selected&" value='"&mydata(myfields("department_id"),rowcounter)&"'>"&mydata(myfields("full_name"),rowcounter)
								Next
								end if
							response.write "</select>"
                            small_help "Dept" %></td>
				</tr>

		
					<tr bgcolor='#FFFFFF'>
						<td width="40%" class="inputname"><b>Plus Customer IDs</b>&nbsp;</td>
						<td width="60%" class="inputvalue"><a href="#" onClick="openCustomerPicker()">Select Customers</a>
							<BR>
							<% if op="edit" then %>
								<textarea rows="4" name="Group_Cid" cols="55"><%= GroupCid %></textarea>
							<% Else %>
								<textarea rows="4" name="Group_Cid" cols="55"></textarea>
							<% end if %>
							<INPUT type="hidden"  name=Group_Cid_C value="Op|String|||||Customer IDs">
                                   <% small_help "Plus Customer ids" %></td>
				</tr>
					

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Group_Company","req","Please enter a company name or enter All to specify all companies.");
 frmvalidator.addValidation("Group_Purchase_History","req","Please enter a purchase minimum.");
 frmvalidator.addValidation("Group_budget_min","req","Please enter a minimum budget.");
 frmvalidator.addValidation("Group_name","req","Please enter a name for this group.");
 frmvalidator.addValidation("group_country","req","Please enter the countries for this group.");

function openCustomerPicker()
{
	window.open("customer_picker.asp","SelectCustomers", "toolbar=no, location=no, width=550, height=600, status=no, scrollbars=yes, resizable=yes");
}
</script>
