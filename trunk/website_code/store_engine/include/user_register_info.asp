<% sub display_register_jstop () %>
<script language="JavaScript" type="text/JavaScript">
function checkBrowser(){
	this.ver=navigator.appVersion
	this.dom=document.getElementById?1:0
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.ns5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie5 || this.ie4 || this.ns4 || this.ns5)
	return this
}
bw=new checkBrowser()
function show(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	obj.display='block'
}
//Hides the div
function hide(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	obj.display='none'
}
function changeState(newLoc){
	if (newLoc.options[newLoc.selectedIndex].value == "United States")
	{
	show('state','State');
	show('state2','State2');
	hide('state_opt','State_Opt')
	hide('state_opt2','State_Opt2')
	hide('state_canada','State_canada')
	hide('state_canada2','State_canada2')
	}
	else
	{
	    if (newLoc.options[newLoc.selectedIndex].value == "Canada")
    	{
		show('state_canada','State_canada')
		show('state_canada2','State_canada2')
		hide('state','State');
		hide('state2','State2');
		hide('state_opt','State_Opt')
		hide('state_opt2','State_Opt2')
    
    	}
    	else
    	{
		hide('state','State');
		hide('state2','State2');
		hide('state_canada','State_canada')
        hide('state_canada2','State_canada2')
		show('state_opt','State_Opt')
		show('state_opt2','State_Opt2')
	    }
	}



}
</script>
<% 
end sub

sub display_register_info ()

if Store_Country = "United States" then
	sDiv="block"
	sDiv1="none"
	sDivCanada="none"
elseif Store_Country="Canada" then
	sDiv="none"
	sDiv1="none"
	sDivCanada="block"
else
	sDiv1="block"
	sDiv="none"
	sDivCanada="none"
end if

%>
<TR>
		<TD class='normal' width='40%'>First Name</TD>
		<TD colspan="2" class='normal'><INPUT value="<%= First_name%>" name=First_name size=30 maxlength=100>
		<INPUT type="hidden"  name=First_name_C value="Re|String|0|100|||First Name">*</TD>
	</TR>

	
	<TR>
		<TD class='normal'>Last Name</TD>
		<TD colspan="2" class='normal'><INPUT value="<%= Last_name%>" name=Last_name size=30 maxlength=100>
		<INPUT type="hidden"  name=Last_name_C value="Re|String|0|100|||Last Name">*</TD>
	</TR>

	<TR>
		<TD class='normal'>Company</TD>
		<TD colspan="2" class='normal'><INPUT value="<%= Company%>" name=company size=30 maxlength=100>
		<INPUT type="hidden"  name=Company_C value="Op|String|0|100|||Company"></TD>
	</TR>

	<TR>
		<TD class='normal'>Address Line 1</TD>
		<TD colspan="2" class='normal'><INPUT value="<%= Address1%>" name=Address1 size=30 maxlength=200>
		<INPUT type="hidden" name=Address1_C value="Re|String|0|200|||Address Line 1">*</TD>
	</TR>
	
	<TR>
		<TD class='normal'>Address Line 2</TD>
		<TD colspan="2" class='normal'><INPUT value="<%= Address2%>" name=address2 size=30 maxlength=200>
		<INPUT type="hidden" name=Address2_C value="Op|String|0|200|||Address Line 2"></TD>
	</TR>

	<TR>
		<TD class='normal'>City</TD>
		<TD colspan="2" class='normal'><INPUT value="<%= City%>" name=City size=30 maxlength=200>
		<INPUT type="hidden"  name=City_C value="Re|String|0|200|||City">*</TD>
	</TR>
        

	<TR>
		<TD class='normal'>Country</TD>
		<TD colspan="2" class='normal'>
			<select size="1" name="Country" OnChange="changeState(this.form.Country)">
			<%  
			sql_country = "SELECT Country,country_id FROM Sys_Countries WITH (NOLOCK) where Country <> 'All Countries' ORDER BY Country;"
			set myfields=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_country,mydata,myfields,noRecords)
			if noRecords = 0 then
			    FOR rowcounter= 0 TO myfields("rowcount")
			        sCountryId=mydata(myfields("country_id"),rowcounter)
			        if Countries_Selected = "ALL" or Is_In_Collection(Countries_Selected,sCountryId,",") then
			            sCountry = mydata(myfields("country"),rowcounter)
			            if sCountry = Store_Country then
				            selected="selected"
			            else
				            selected=""
			            end if 
			            response.write "<Option "&selected&" value="""&sCountry&""">"&sCountry&"</option>"
			        end if
		        Next 
			End If 
%>
			</select>*</TD>
	</TR>



	<tr>
	<td class='normal'>
	<DIV NAME="state2" ID="state2" style="display: <%= sDiv %>;">State</div>
	<DIV NAME="state_canada2" ID="state_canada2" style="display: <%= sDivCanada %>;">Province</div>
    <DIV NAME="state_opt2" ID="state_opt2" style="display: <%= sDiv1 %>;">County/Region/Other</div></td>
	<td class='normal'>
  <DIV NAME="state" ID="state" style="display: <%= sDiv %>;">
  <select name='State_UnitedStates'>
	<option value=''>Please select</option>
	<% 
	sql_sel = "select State,State_Long,Country from Sys_States WITH (NOLOCK) order by state_long"
	set statefields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_sel,statedata,statefields,noRecords)
	if noRecords = 0 then
	    FOR staterowcounter= 0 TO statefields("rowcount")
	        sCountry = statedata(statefields("country"),staterowcounter)
	        if sCountry="United States" then
	            sState = statedata(statefields("state"),staterowcounter)
	            sStateLong = statedata(statefields("state_long"),staterowcounter)
                if sState = State then
	                selected="selected"
                else
	                selected=""
                end if 
		        response.write "<option "&selected&" value='"&sState&"'>"&statedata(statefields("state_long"),staterowcounter)&"</option>"
	        end if
	    next
	End If %> 
	</select>*
    </div>
	<DIV NAME="state_canada" ID="state_canada" style="display: <%= sDivCanada %>;">
	<select name='State_Canada'>
	<option value=''>Please select</option>
    <% if noRecords = 0 then
	    FOR staterowcounter= 0 TO statefields("rowcount")
	        sCountry = statedata(statefields("country"),staterowcounter)
	        if sCountry="Canada" then
	            sState = statedata(statefields("state"),staterowcounter)
	            sStateLong = statedata(statefields("state_long"),staterowcounter)
                if sState = State then
	                selected="selected"
                else
	                selected=""
                end if 
		        response.write "<option "&selected&" value='"&sState&"'>"&statedata(statefields("state_long"),staterowcounter)&"</option>"
	        end if
	    next
	End If
	%>
  </select>*
    </div>
    <DIV NAME="state_opt" ID="state_opt" style="display: <%= sDiv1 %>;">
    <INPUT name=State_Opt value="<%= State %>" size=30 maxlength=50>
    <INPUT type="hidden"  name=State_Opt_C value="Op|String|0|50|||State">
    </div></td>
</tr>

	<TR>
		<TD class='normal'>Zip/Postal Code</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Zip size="10" maxlength=14 value="<%= Zip %>">
			<INPUT type="hidden"  name=zip_C value="Re|String|0|14|||Zip Code">*</TD>
	</TR>

	<TR>
		<TD class='normal'>Phone</TD>
		<TD colspan="2" class='normal'>
			<INPUT value="<%= Phone%>" name=phone size=30 maxlength=50>
			<INPUT type="hidden"  name=phone_C value="Re|String|0|50|||Phone">*</TD>
	</TR>

	<TR>
		<TD class='normal'>Email</TD>
		<TD colspan="2" class='normal'>
			<INPUT value="<%= Email%>" name=Email size=30 maxlength=100 onKeyPress="return goodchars(event,'abcdefghijklmnopqrstuvwxyz@0123456789.-_+*!#$%&/=?^{}|~')">
			<INPUT type="hidden"  name=Email_C value="Re|Email|0|100|||Email">*</TD>
	</TR>

	<TR>
		<TD class='normal'>Fax</TD>
		<TD colspan="2" class='normal'>
			<INPUT value="<%= Fax%>" name=Fax size=30 maxlength=50>
			<INPUT type="hidden"  name=Fax_C value="Op|String|0|50|||Fax"></TD>
	</TR>

<% end sub
sub display_register_jsbottom () %>
frmvalidator.addValidation("First_name","req","Please enter a first name.");
 frmvalidator.addValidation("Last_name","req","Please enter a last name.");
 frmvalidator.addValidation("Address1","req","Please enter a address.");
 frmvalidator.addValidation("City","req","Please enter a city.");
 frmvalidator.addValidation("Country","req","Please enter a country.");
 frmvalidator.addValidation("Zip","req","Please enter a postal code.");
 frmvalidator.addValidation("phone","req","Please enter a phone number.");
 frmvalidator.addValidation("Email","req","Please enter a email.");
 frmvalidator.addValidation("Email","email","Please enter a valid email.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation2");
 function DoCustomValidation2()
{
  var frm = document.forms["<%= sFormName %>"];

	if (frm.State_UnitedStates.options[frm.State_UnitedStates.selectedIndex].value == "" && frm.Country.options[frm.Country.selectedIndex].value == "United States")
		{ 
		alert('Please choose a state.');
		return false
		}
	else if (frm.State_Canada.value == "" && frm.Country.options[frm.Country.selectedIndex].value == "Canada")
		{
		  alert('Please choose a province.');
		  return false
		}
  return true

}
<% end sub %>
