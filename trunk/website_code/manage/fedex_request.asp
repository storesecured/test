<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/country_list.asp"-->


<%

sFormAction = "http://shippingserver2.storesecured.com/fedexapigetmeter.aspx"

sTitle = "Fedex Meter Number"
thisRedirect = "fedex_request.asp"
sMenu="shipping"
createHead thisRedirect  %>


      
<TR bgcolor='#FFFFFF'><td colspan=3><B>Please fill out the information below to receive your FedEx Meter no</td></tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">FedEx Account Number</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="Account_No" value="" size="60" maxlength=200>* Only Numbers
		<input type="hidden" name="Account_No_C" value="Re|Integer|0|999999999|||Account Number">
		<% small_help "Account Number" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">FedEx Account Owner</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="Name" value="" size="60" maxlength=200>* Your name
		<input type="hidden" name="Name_C" value="Re|String|0|200|||Account Owner">
		<% small_help "Account Owner" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">Phone Number</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="Phone" value="<%= replace(replace(replace(replace(Store_Phone,"-",""),"(",""),")","")," ","") %>" size="60" maxlength=200>* Only numbers
		<input type="hidden" name="Phone_C" value="Re|Integer|||||Phone">
		<% small_help "Phone" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">Address</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="Address" value="<%= Store_Address1 %>" size="60" maxlength=200>*
		<input type="hidden" name="Address_C" value="Re|String|0|200|||Address">
		<% small_help "Address" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">City</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="City" value="<%= Store_City %>" size="60" maxlength=200>*
		<input type="hidden" name="City_C" value="Re|String|0|200|||City">
		<% small_help "Address" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">State Code</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="State" value="<%= UCase(Store_State) %>" size="60" maxlength=2>* 2 Char Abbreviation
		<input type="hidden" name="State_C" value="Re|String|0|2|||State">
		<% small_help "Address" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">Postal Code</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="Zip" value="<%= Store_Zip %>" size="60" maxlength=200>*
		<input type="hidden" name="Zip_C" value="Re|String|0|200|||Zip">
		<% small_help "Address" %></td>
</tr>
<tr bgcolor=#FFFFFF>
        <td width="24%" height="23" class="inputname">Country Code</td>
	<td width="76%" height="23" class="inputvalue">
		<input type="text" name="Country" value="US" size="60" maxlength=2>* 2 Char Abbreviation
		<input type="hidden" name="Country_C" value="Re|String|0|2|||Country">
		<% small_help "Country" %></td>
</tr>


<% createFoot thisRedirect, 1 %>


