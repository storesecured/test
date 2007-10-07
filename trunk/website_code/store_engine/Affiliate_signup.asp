<!--#include file="include/header.asp"-->
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
	hide('state_opt','State_Opt')

	}
	else
	{
	    if (newLoc.options[newLoc.selectedIndex].value == "Canada")
    	{
      show('state','State');
    	hide('state_opt','State_Opt')
    
    	}
    	else
    	{
      hide('state','State');
	    show('state_opt','State_Opt')
	    }
	}



}
</script>
<%
if Store_Country = "United States" or Store_Country="Canada" then
	 sDiv="block"
	 sDiv1="none"
else
	 sDiv1="block"
	 sDiv="none"
end if
%>
<form method="POST" action="<%= Site_Name %>Affiliates_Action.asp" name="Registration">
<input type="Hidden" name="Record_type" value="0">
<TABLE WIDTH="311" BORDER=0 CELLSPACING=1 CELLPADDING=1 style="height: 114px">

		<TR>
			<TD	colspan=3 width="274"><b>Create Affiliate Account</b></TD>
		</TR>
		
		<TR>
			<TD width="131">Username</font></TD>
			<TD width="164" colspan="2">
				<INPUT  name=Code size="30">*
				<INPUT type="hidden"  name=Code_C value="Re|String|0|50|||Username"></TD>
		</TR>
		
		<TR> 
			<TD width="131">Password</TD>
			<TD width="164" colspan="2">
				<INPUT name=Password type=password size="30">*
				<INPUT name=Password_C type=hidden value="Re|String|0|50|||Password" size="30"></TD>
		</TR>
		
		<TR>
			<TD width="131">Confirm Password</TD>
			<TD width="164" colspan="2">
				<INPUT name=Password_Confirm type=password size="30">*
				<INPUT name=Password_Confirm_C type=hidden value="Re|String|0|50|||Confirm Password" size="30"></TD>
		</TR>
	
	<TR>
		<TD width="131">Contact Name</TD>
		<TD width="164" colspan="2">
		<INPUT  name=Contact_Name size="30">*
		<INPUT type="hidden"  name=Contact_Name_C value="Re|String|0|50|||Contact Name"></TD>
	</TR>
	
	
	<TR>
		<TD width="131">Email</TD>
		<TD width="164" colspan="2">
		<INPUT  name=Email size="30">*
		<INPUT type="hidden"  name=Email_C value="Re|Email|0|50|||Email"></TD>
	</TR>

	<TR>
		<TD width="131">URL</TD>
		<TD width="164" colspan="2">
		<INPUT  name=URL size="30" value="http://">
		<INPUT type="hidden"  name=URL_C value="Re|String|0|255|http://||URL"></TD>
	</TR>

   <TR>
		<TD class='normal'>Company</TD>
		<TD colspan="2" class='normal'><INPUT	name=company size="30">
		<INPUT type="hidden"  name=Company_C value="Op|String|0|100|||Company"></TD>
	</TR>

	<TR>
		<TD class='normal'>Address</TD>
		<TD colspan="2" class='normal'><INPUT	name=Address size="30">*
		<INPUT type="hidden" name=Address_C value="Re|String|0|200|||Address"></TD>
	</TR>

	<TR>
		<TD class='normal'>City</TD>
		<TD colspan="2" class='normal'><INPUT	name=City size="30">*
		<INPUT type="hidden"  name=City_C value="Re|String|0|200|||City"></TD>
	</TR>

	<TR>
		<TD class='normal'>Country</TD>
		<TD colspan="2" class='normal'>
			<select size="1" name="Country" OnChange="changeState(this.form.Country)">
				<% sql_country = "SELECT Country FROM Sys_Countries WITH (NOLOCK) where Country <> 'All Countries' ORDER BY Country;"
				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_country,mydata,myfields,noRecords)


				if noRecords = 0 then
					FOR rowcounter= 0 TO myfields("rowcount")
					if mydata(myfields("country"),rowcounter) = Store_Country then
						selected="selected"
					else
						selected=""
					end if %>
					<Option <%= selected %> value="<%= mydata(myfields("country"),rowcounter) %>"> <%= mydata(myfields("country"),rowcounter) %></option>
					<% Next %>
					<% End If %>
			</select>*</TD>
	</TR>
	<tr>
	<td class='normal'>State</td>
	<td class='normal'>
		<DIV NAME="state" ID="state" style="display: <%= sDiv %>;">
		<select name='State'>
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
    <DIV NAME="state_opt" ID="state_opt" style="display: <%= sDiv1 %>;">
    <INPUT name=State_Opt value="" size=20 maxlength=50>*
    <INPUT type="hidden"  name=State_Opt_C value="Op|String|0|50|||State">
    </div></td>
</tr>

	<TR>
		<TD class='normal'>Zip Code</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Zip size="10">*
			<INPUT type="hidden"  name=zip_C value="Re|String|0|14|||Zip Code"></TD>
	</TR>
	
	<TR>
		<TD class='normal'>Phone</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=phone size="30">*
			<INPUT type="hidden"  name=phone_C value="Re|String|0|50|||Phone"></TD>
	</TR>

	<TR>

		<TD width="97">
			<%= fn_create_action_button ("Button_image_Register", "Register_User", "Register") %>
			</TD>
	</TR>
</TABLE>

</form>

<!--#include file="include/footer.asp"-->
