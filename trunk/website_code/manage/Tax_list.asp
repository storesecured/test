<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

'LOAD TAX RATES LIST FROM THE DATABASE
sql_taxes = "SELECT Zip_Name, Zip_Start, Zip_End, TaxRate, Zip_Range_ID, Department_Ids, Tax_Shipping FROM Store_Zips WHERE Store_id = "&Store_id
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_taxes,mydata,myfields,noRecords)

sql_dept = "SELECT department_name, department_Id FROM store_dept where store_id="&store_id&" ORDER BY department_name;"
set myfields1=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_dept,mydata1,myfields1,noRecords1)

sFlashHelp="taxes/taxes.htm"
sMediaHelp="taxes/taxes.wmv"
sZipHelp="taxes/taxes.zip"
sTextHelp="taxes/general-tax-list.doc"

sInstructions="The customers ship to address will be used to determine their tax rate.	If multiple tax rates apply they will be added together to determine the final tax rate.  Areas for which no taxes are defined will be set to 0%."
	
sFormAction = "update_records.asp"
sName = "tax"
sTitle = "Taxes"
sFullTitle = "General > Taxes"
thisRedirect = "tax_list.asp"
sMenu = "general"
sQuestion_Path = "general/taxes.htm"
createHead thisRedirect
%>

	
	 <tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="0">
			<input type="button" class="Buttons" value="Add State Tax" name="Add_Tax" OnClick=JavaScript:self.location="tax_add_state.asp?<%= sAddString %>"><input type="button" class="Buttons" value="Add Zipcode Tax" name="Add_Tax" OnClick=JavaScript:self.location="tax_add.asp?<%= sAddString %>"><input type="button" class="Buttons" value="Add Country Tax" name="Add_Tax" OnClick=JavaScript:self.location="tax_add_country.asp?<%= sAddString %>"></td>
	 </tr>
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3">

				
				<% if noRecords = 0 then %>
				<table width="100%" border="0" cellpadding="2" cellspacing="0" class="list">

				<tr bgcolor='#0069B2' class='white'>
				<td height="4" class=tablehead width="50%"><font class=white><B>Name</B></font></td>
				<td height="4" class=tablehead><font class=white><B>Zip Range</B></font></td>

				<td height="4" class=tablehead><font class=white><B>Tax Rate</B></font></td>
					<td height="4" class=tablehead><font class=white><B>Departments</B></font></td>
					<td height="4" class=tablehead><font class=white><B>Tax Shipping</B></font></td>
				<td height="4" class=tablehead>&nbsp;</td>
				</tr> 
				<%
				strclass=1
				FOR rowcounter= 0 TO myfields("rowcount")
					if str_class=1 then
						str_class=0
						bgcolor="#FFFFFF"
					else
						str_class=1
						bgcolor="#ECEFF0"
					end if
					response.write "<tr class='"&str_class&"' bgcolor='"&bgcolor&"'>" %>
						<td width="10%" height="4" class=<%=str_class%>><%= mydata(myfields("zip_name"),rowcounter) %></td>
						<td width="60%" height="4" class=<%=str_class%>><%= Zipcode_Format_Function(mydata(myfields("zip_start"),rowcounter)) %> - <%= Zipcode_Format_Function(mydata(myfields("zip_end"),rowcounter)) %></td>
						<td width="10%" height="4" class=<%=str_class%>><%= formatNumber(mydata(myfields("taxrate"),rowcounter),4) %>%</td>
						<td width="10%" height="4" class=<%=str_class%>>
							<select name="Departments" size="1">
								<%
								Collection = Cstr(mydata(myfields("department_ids"),rowcounter))
								If Is_In_Collection(Collection,"-1",",") then
									response.write "<Option value='-1'>All Depts</option>"
								end if
								if noRecords1 = 0 then
									FOR rowcounter1= 0 TO myfields1("rowcount")
										one_item = Cstr(mydata1(myfields1("department_id"),rowcounter1))
										' diplay only those countries that are inside our Countries list
										' let's use this function :	Is_In_Collection(Collection,one_item,Del)
										if Is_In_Collection(Collection,one_item,",") then
											response.write "<Option value='"&mydata1(myfields1("department_id"),rowcounter1)&"'>"&mydata1(myfields1("department_name"),rowcounter1)&"</option>"
										end if
									Next
								End if %>
							</select>
							</td>
							<td class=<%=str_class%>>
							<% if mydata(myfields("tax_shipping"),rowcounter) then
								Tax_Shipping="Yes"
								else
								Tax_Shipping="No"
								end if
								response.write Tax_Shipping
								%>

							</td>
					<td width="36%" height="4" nowrap class=<%=str_class%>>
						<a class="link" href="tax_add.asp?op=edit&Zip_Range_ID=<%= mydata(myfields("zip_range_id"),rowcounter) %>&<%= sAddString %>">Edit</a>&nbsp;
							<a class="link" href="tax_action.asp?Delete_Zip_Range_ID=<%= mydata(myfields("zip_range_id"),rowcounter) %>&<%= sAddString %>" OnClick="return VerifyDelete(0);">Delete</a></td>
					</tr>
				<% Next %>
				</table>
				<BR>
				<% End if %>


		<% sql_taxes = "SELECT State, TaxRate, State_Tax_ID, Department_Ids, Tax_Shipping FROM Store_State_Tax WHERE Store_id = "&Store_id
		  set myfields=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_taxes,mydata,myfields,noRecords)
			 %>



				<% if noRecords = 0 then %>
				<table width="100%" border="0" cellpadding="2" cellspacing="0" class="list">

				<tr bgcolor='#0069B2' class='white'>
				<td width="10%" height="4" class=tablehead><font class=white><B>States</B></font></td>

				<td width="10%" height="4" class=tablehead><font class=white><B>Tax Rate</B></font></td>
					<td width="20%" height="4" class=tablehead><font class=white><B>Departments</B></font></td>
				<td class=tablehead><font class=white><B>Tax Shipping</b></font></td>
				<td width="10%" height="4" class=tablehead>&nbsp;</td>
				</tr>
				<%
				strclass=1
				FOR rowcounter= 0 TO myfields("rowcount")
					 if str_class=1 then
						str_class=0
						bgcolor="#FFFFFF"
					else
						str_class=1
						bgcolor="#ECEFF0"
					end if
					response.write "<tr class='"&str_class&"' bgcolor='"&bgcolor&"'>" %>
					
						<td width="10%" height="4" class=<%=str_class%>><select name="State" size="1">
								<% StateArray=split(mydata(myfields("state"),rowcounter),",") %>
								<% For i=0	to ubound(StateArray)  %>
									<% State=StateArray(i) %>
									<Option value="<%= State %>"> <%= State %></option>

								<% Next %>
							</select></td>
						<td width="10%" height="4" class=<%=str_class%>><%= formatNumber(mydata(myfields("taxrate"),rowcounter),4) %>%</td>
						<td width="10%" height="4" class=<%=str_class%>>
							<select name="Departments" size="1">
								<% Collection = Cstr(mydata(myfields("department_ids"),rowcounter))
								If Is_In_Collection(Collection,"-1",",") then
									response.write "<Option value='-1'>All Depts</option>"
								end if
								if noRecords1 = 0 then
									FOR rowcounter1= 0 TO myfields1("rowcount")
										one_item = Cstr(mydata1(myfields1("department_id"),rowcounter1))
										' diplay only those countries that are inside our Countries list
										' let's use this function :	Is_In_Collection(Collection,one_item,Del)
										if Is_In_Collection(Collection,one_item,",") then
											response.write "<Option value='"&mydata1(myfields1("department_id"),rowcounter1)&"'>"&mydata1(myfields1("department_name"),rowcounter1)&"</option>"
										end if
									Next
								End if %>
							</select>
							</td>
							<td class=<%=str_class%>>
							<% if mydata(myfields("tax_shipping"),rowcounter) then
								Tax_Shipping="Yes"
								else
								Tax_Shipping="No"
								end if
								response.write Tax_Shipping
								%>

							</td>
					<td width="36%" height="4" nowrap class=<%=str_class%>>
						<a class="link" href="tax_add_state.asp?op=edit&State_Tax_Id=<%= mydata(myfields("state_tax_id"),rowcounter) %>&<%= sAddString %>">Edit</a>&nbsp;
							<a class="link" href="tax_action.asp?Delete_State_Tax_ID=<%= mydata(myfields("state_tax_id"),rowcounter) %>&<%= sAddString %>" OnClick="return VerifyDelete(0);">Delete</a></td>
					</tr>
				<% Next %>
				</table>
				<BR>
				<% End If %>



		<% sql_taxes_country = "SELECT Country, TaxRate, Country_Tax_ID, Department_Ids, Tax_Shipping FROM Store_Country_Tax WHERE Store_id = "&Store_id
		  set myfields=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_taxes_country,mydata,myfields,noRecords) %>



				<% str_class=1
				if noRecords = 0 then %>
				<table width="100%" border="0" cellpadding="2" cellspacing="0" class="list">

				<tr bgcolor='#0069B2' class='white'>
				<td width="10%" height="4" class=tablehead><font class=white><B>Countries</B></font></td>

				<td width="20%" height="4" class=tablehead><font class=white><B>Tax Rate</B></font></td>
					<td width="20%" height="4" class=tablehead><font class=white><B>Departments</B></font></td>
				<td class=tablehead><font class=white><B>Tax Shipping</b></font></td>
				<td width="10%" height="4" class=tablehead>&nbsp;</td>
				</tr>
				<%
				str_class=1
				FOR rowcounter= 0 TO myfields("rowcount")
					if str_class=1 then
						str_class=0
						bgcolor="#FFFFFF"
					else
						str_class=1
						bgcolor="#ECEFF0"
					end if
					response.write "<tr class='"&str_class&"' bgcolor='"&bgcolor&"'>" %>
						<td width="10%" height="4" class=<%=str_class%>><select name="Country" size="1">
								<% CountryArray=split(mydata(myfields("country"),rowcounter),",") %>
								<% For i=0	to ubound(CountryArray)  %>
									<% Country=CountryArray(i) %>
									<Option value="<%= Country %>"> <%= Country %></option>

								<% Next %>
							</select></td>
						<td width="10%" height="4" class=<%=str_class%>><%= formatNumber(mydata(myfields("taxrate"),rowcounter),4) %>%</td>
						<td width="10%" height="4" class=<%=str_class%>>
							<select name="Departments" size="1">
								<% Collection = Cstr(mydata(myfields("department_ids"),rowcounter))
								If Is_In_Collection(Collection,"-1",",") then
									response.write "<Option value='-1'>All Depts</option>"
								end if
								if noRecords1 = 0 then
									FOR rowcounter1= 0 TO myfields1("rowcount")
										one_item = Cstr(mydata1(myfields1("department_id"),rowcounter1))
										' diplay only those countries that are inside our Countries list
										' let's use this function :	Is_In_Collection(Collection,one_item,Del)
										if Is_In_Collection(Collection,one_item,",") then
											response.write "<Option value='"&mydata1(myfields1("department_id"),rowcounter1)&"'>"&mydata1(myfields1("department_name"),rowcounter1)&"</option>"
										end if
									Next
								End if %>
							</select>
							</td>
							<td class=<%=str_class%>>
							<% if mydata(myfields("tax_shipping"),rowcounter) then
								Tax_Shipping="Yes"
								else
								Tax_Shipping="No"
								end if
								response.write Tax_Shipping
								%>

							</td>
					<td width="36%" height="4" nowrap class=<%=str_class%>>
						<a class="link" href="tax_add_country.asp?op=edit&Country_Tax_Id=<%= mydata(myfields("country_tax_id"),rowcounter) %>&<%= sAddString %>">Edit</a>&nbsp;
							<a class="link" href="tax_action.asp?Delete_Country_Tax_ID=<%= mydata(myfields("country_tax_id"),rowcounter) %>&<%= sAddString %>"  OnClick="return VerifyDelete(0);">Delete</a></td>
					</tr>
				<% Next %>
				</table>
				<% End If %>

		</td>
	</tr>
<% createFoot thisRedirect, 0%>
