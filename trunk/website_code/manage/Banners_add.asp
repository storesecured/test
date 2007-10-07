<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 
sTextHelp="banners/banners.doc"

calendar=1
BID = "-1"
Banner_Name = ""
Image_Path = ""
URL = "http://"
Start_Date = dateserial(year(now()), month(now()), day(now()))
End_Date = dateadd("m",1,Start_Date)
RunDays = vbMonday&", "&vbTuesday&", "&vbWednesday&", "&vbThursday&", "&vbFriday&", "&vbSaturday&", "&vbSunday
StartHour = 0
EndHour = 24
BID = request.querystring("ID")
if BID = "" then
   BID = request.querystring("BID")
end if

if BID<>"" then
	if not isNumeric(BID) then
	   Response.Redirect "admin_error.asp?message_id=1"
	end if
	sql_sel = "select * from store_banners where store_id="&store_id&" and bann_id="&bid
	rs_store.open sql_sel, conn_store, 1, 1
	     if rs_store.bof or rs_store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	     end if
		Banner_Name = rs_store("Banner_Name")
		Image_Path = rs_store("Image_Path")
		sURL = rs_store("URL")
		Start_Date = rs_store("Start_Date")
		End_Date = rs_store("End_Date")
		RunDays = rs_store("RunDays")
		StartHour = rs_store("StartHour")
		EndHour = rs_store("EndHour")
		Enabled = rs_store("Enabled")
	rs_store.close

	if Enabled then
		Enabledchecked = "checked"
	else
		Enabledchecked = ""
	end if
end if

sCommonName="Banner"
sCancel="banners.asp"
If BID="" then
   sTitle = "Add Banner"
   sFullTitle = "Marketing > <a href=banners.asp class=white>Banners</a> > Add"
Else
   sTitle = "Edit Banner - "&Banner_Name
   sFullTitle = "Marketing > <a href=banners.asp class=white>Banners</a> > Edit - "&Banner_Name
End If

addPicker = 1
sFormAction = "Banners_action.asp"
thisRedirect = "banners_add.asp"
sMenu = "marketing"
sQuestion_Path = "marketing/banners.htm"
createHead thisRedirect
%>




<% If BID="" then %>
	<input type="hidden" name="Action" value="Add">
<% Else %>
	<input type="hidden" name="Action" value="Edit">
	<input type="hidden" name="BID" value="<%= BID %>">
<% End If %>
				
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Enabled</b></td>
					<td width="90%" class="inputvalue"><input class="image" type="checkbox" name="Enabled" <%= Enabledchecked %> size="60">
					<% small_help "Enabled" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Name</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=50 name="Banner_Name" value="<%= Banner_Name %>" size="60">
					<INPUT type="hidden"  name="Banner_Name_C" value="Re|String|0|50|||Name">
					<% small_help "Name" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Image</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=45 name="Image_Path" value="<%= Image_Path %>" size="60"><a href="JavaScript:goImagePicker('Image_Path');"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
					<INPUT type="hidden"  name="Image_C" value="Op|String|0|45|||Image">
					<a class="link" href="JavaScript:goFileUploader('Image_Path');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
<% small_help "Image" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>URL</b></td>
					<td width="90%" class="inputvalue"><input type="text" maxlength=100 name="Url" value="<%= sURL %>" size="60">
					<INPUT type="hidden"  name="Url_C" value="Re|String|0|100|||URL">
					<% small_help "URL" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Days</b></td>
					<td width="90%" class="inputvalue">
						<select multiple name="RunDays">
							<option value="<%= vbMonday %>"
								<% If instr(RunDays&",", vbMonday&",")>0 then %>selected<% End If %>>Monday</option>
							<option value="<%= vbTuesday %>"
								<% If instr(RunDays&",", vbTuesday&",")>0 then %>selected<% End If %>>Tuesday</option>
							<option value="<%= vbWednesday %>"
								<% If instr(RunDays&",", vbWednesday&",")>0 then %>selected<% End If %>>Wednesday</option>
							<option value="<%= vbThursday %>"
								<% If instr(RunDays&",", vbThursday&",")>0 then %>selected<% End If %>>Thursday</option>
							<option value="<%= vbFriday %>"
								<% If instr(RunDays&",", vbFriday&",")>0 then %>selected<% End If %>>Friday</option>
							<option value="<%= vbSaturday %>"
								<% If instr(RunDays&",", vbSaturday&",")>0 then %>selected<% End If %>>Saturday</option>
							<option value="<%= vbSunday %>"
								<% If instr(RunDays&",", vbSunday&",")>0 then %>selected<% End If %>>Sunday</option>
						</select>
						<% small_help "Days" %>
					</td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Hours</b></td>
					<td width="90%" class="inputvalue">
						<select name="StartHour">
							<% for i=0 to 24 %>
								<% if i<10 then %>
									<option value="<%= i %>"
										<% If cstr(i)=cstr(StartHour) then %>selected<% End If %>
										>0<%= i %>:00</option>
								<% Else %>
									<option value="<%= i %>"
										<% If cstr(i)=cstr(StartHour) then %>selected<% End If %>><%= i %>:00</option>
								<% End If %>
							<% next %>
						</select>
						&nbsp;-&nbsp;
						<select name="EndHour">
							<% for i=0 to 24 %>
								<% if i<10 then %>
									<option value="<%= i %>"
										<% If cstr(i)=cstr(EndHour) then %>selected<% End If %>
										>0<%= i %>:00</option>
								<% Else %>
									<option value="<%= i %>"
										<% If cstr(i)=cstr(EndHour) then %>selected<% End If %>
										><%= i %>:00</option>
								<% End If %>
							<% next %>
						</select>
						<% small_help "Hours" %></td>
				</tr><SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>Start Date</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="Start_Date" value="<%= Start_Date %>" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
					<A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
			   <INPUT type="hidden"  name="Start_Date_C" value="Op|date|||||Start Date">
					<% small_help "Start Date" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td width="10%" class="inputname"><b>End Date</b></td>
					<td width="90%" class="inputvalue"><input type="text" name="End_Date" value="<%= End_Date %>" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
					<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
        <INPUT type="hidden"  name="End_Date_C" value="Op|date|||||End Date">
					<% small_help "End Date" %></td>
				</tr>

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Banner_Name","req","Please enter a name for this banner.");
 frmvalidator.addValidation("Url","req","Please enter a destination url.");
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");

</script>
