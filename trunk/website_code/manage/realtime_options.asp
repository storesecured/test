<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


	' IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlSelShipping="select * from Store_real_time_settings where " & _
						"Store_id="&Store_id
	rs_store.open sqlSelShipping,conn_store,1,1
	Countries=rs_store("Countries")
	Restrict_Options=rs_store("Restrict_Options")
	Max_Weight=rs_Store("Max_Weight")
	rs_store.close


sFormAction = "update_records_action.asp"
sName = "realtime"
sTitle = "Realtime Shipping Options"
thisRedirect = "realtime_options.asp"
sMenu = "general"
sQuestion_Path = "general/shipping.htm"
createHead thisRedirect


%>
<input type=hidden name=realtime value=1>
          <tr>
					<td width="20%"class="inputname"><B>Max Weight Per Box</b></td>
					<td width="80%" class="inputvalue">
							<input type=text name="Max_Weight" onKeyPress="return goodchars(event,'0123456789.')" value=<%= Max_Weight %>>
						<input type="hidden" name="Max_Weight_C" value="Re|Integer|0|150|||Max Weight">
						<% small_help "Max_Weight" %></td>
					</tr>
          <tr>
					<td width="20%"class="inputname"><B>Restricted Options</b></td>
					<td width="80%" class="inputvalue">
							<textarea name="Restrict_Options" rows=5 cols=40><%= Restrict_Options %></textarea>
						<input type="hidden" name="Restrict_Options_C" value="Re|String|0|1000|||Restricted Options">
						<% small_help "Name" %></td>
					</tr>

					<tr>
					<td align="left" width="188" height="20" class="inputname"><B>Allowed Countries</b></td>
					<td align="left" width="186" height="20" class="inputvalue">
						<select multiple name="Countries" size="6">
							<Option value="">Select valid countries for this shipping method
							<%
							sql_region = "SELECT Country,Country_Id FROM Sys_Countries ORDER BY Country;"
							set myfields1=server.createobject("scripting.dictionary")
							Call DataGetrows(conn_store,sql_region,mydata1,myfields1,noRecords1)
							if noRecords1 = 0 then
							FOR rowcounter1= 0 TO myfields1("rowcount")
								' set the selected flag

									CountriesArray=split(Countries,",")
									Found=false
									For i=0 to ubound(CountriesArray)
										CountryElem=replace(CountriesArray(i)," ","")
										Country=replace(mydata1(myfields1("country"),rowcounter1)," ","")
										If CountryElem = Country  Then
											Found=true
											exit for
										End If
									Next
									if Found then
										selected = "selected"
									else
										selected=""
									end if

								response.write "<Option value='"&mydata1(myfields1("country"),rowcounter1)&"' "&selected&" >"&mydata1(myfields1("country"),rowcounter1)&"</option>"
							Next
							End If
							%>
						</select>
						<input type="hidden" name="Countries_C" value="Re|String|0|4000|||Countries">
						<% small_help "Countries" %></td>
					</tr>

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Countries","req","Please select the countries that realtime shipping will apply to.");
 frmvalidator.addValidation("Max_Weight","req","Please enter a maximum weight.");

</script>

