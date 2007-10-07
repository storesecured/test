<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
objHelpDict.Add "force_resend","Bypass regular weekly or monthly send settings to create a feed tonight."

sql_select = "SELECT * FROM Store_External WHERE Store_id="&Store_id
rs_Store.open sql_select,conn_store,1,1
rs_Store.MoveFirst

Froogle_Enable = rs_Store("Froggle_Enable")
Froogle_Filename = rs_Store("Froggle_Filename")
Froogle_Username = rs_Store("Froggle_Username")
Froogle_Password = rs_Store("Froggle_Password")
Froogle_Server = rs_Store("Froggle_Server")
Froogle_Submit = rs_Store("Froggle_Submit")
Large_Upload = rs_Store("Large_Upload")
Force_Resend = rs_Store("Force_Resend")
if Froogle_Enable then
	froogle_checked = "checked"
else
	froogle_checked = ""
end if
if Force_Resend then
	force_checked = "checked"
else
	force_checked = ""
end if
if Large_Upload then
	large_checked = "checked"
	weekly_checked=""
else
	large_checked = ""
	weekly_checked="checked"
end if
rs_Store.close

if Froogle_Submit = "" then
	Froogle_Submit = "Never"
end if
sFormAction = "Store_Settings.asp"
sName = "Store_Froogle"
sFormName = "Froogle"
sTitle = "Google Base"
sFullTitle = "Marketing > Google Base"
sSubmitName = "Update_Froogle"
thisRedirect = "froogle_settings.asp"
sTopic = "Froogle"
AddPicker=1
sMenu = "marketing"
sQuestion_Path = "marketing/froogle.htm"
createHead thisRedirect

%>
  <tr bgcolor='#FFFFFF'>
  <td width="100%" colspan="5" class=instructions>Google Base (formerly Froogle) is an extension of the Google search engine that
  millions of people around the globe use daily to research products before they purchase. Listing
  your product in Google Base is a free way to extend the reach of your marketing efforts to millions of new customers.
  Google's worldwide user base performs more than 150 million searches a day. That presents an enormous
  opportunity to introduce your products to customers you might otherwise spend millions of dollars in
  advertising to reach. These are people actively searching for the items you sell. And did we mention
  inclusion in Google Base is absolutely free?<BR><BR>
  Sign up with Google Base and enter your username and password and we will automatically send them your updated
  product list once a week.   


  </td>
  </tr>
  <tr bgcolor='#FFFFFF'>
  <td width="100%" colspan="5"><A class="link" HREF="http://services.google.com/froogle/merchant_email" target="_blank">Learn more about <b>Google Base</b> and apply for an account</A>
  </td>
  </tr>
  
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Enable Google Base Feed</B>
		</td><td width="75%" class="inputvalue"><input class="image" type="checkbox" <%= froogle_checked %> name="Froogle_Enable">
	  <% small_help "Enable Feed" %></td>


  </tr>
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Google Base FTP Username</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= Froogle_Username %>" name="Froogle_Username" size="25" maxlength=50>
		  <INPUT type="hidden"	name="Froogle_Username_C" value="Re|String|0|50|||Froogle Username">
		 <% small_help "Username" %></td>
  </td>
  </tr>
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Google Base FTP Password</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= Froogle_Password %>" name="Froogle_Password" size="25" maxlength=50>
		  <INPUT type="hidden"	name="Froogle_Password_C" value="Re|String|0|50|||Froogle Password">
		 <% small_help "Password" %></td>
  </td>
  </tr>
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Google Base FTP Filename</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= Froogle_Filename %>" name="Froogle_Filename" size="25" maxlength=50>.txt
		  <INPUT type="hidden"	name="Froogle_Filename_C" value="Re|String|0|50|||Froogle Filename">
		 <% small_help "Filename" %></td>
  </td>
  </tr>
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Google Base FTP Server Address</B></td>
  <td width="75%" class="inputvalue" nowrap>
		  ftp://<input type="text" value="<%= Froogle_Server %>" name="Froogle_Server" size="25" maxlength=50>
		  <INPUT type="hidden"	name="Froogle_Server_C" value="Re|String|0|50|google.com||Froogle Server">
		 <% small_help "Server Address" %></td>
  </td>
  </tr>

  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Send Frequency</B></td>
  <td width="75%" class="inputvalue" nowrap>
		  <input class="image" type="radio" value="1" <%= large_checked %> name="Large_Upload"> Monthly
		  <BR><input class="image" type="radio" value="0" <%= weekly_checked %> name="Large_Upload"> Weekly
		  <BR>Weekly feeds are capped at 3000 items maximum <% small_help "Large_Upload" %></td>
  </td>
  </tr>
  
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname" height=20><B>Last Successfull Feed</B></td>
  <td width="75%" class="inputvalue">
		  <% if Froogle_Submit="" or isNull(Froogle_Submit) then %>
			Your feed has never been sent successfully.  If your feed has been enabled for more then 24 hours please check your username and password entered above.
		  <% else %>
		  <%= Froogle_Submit %>
		  <% end if %>
		  <% small_help "Last Feed" %></td>
  </td>
  </tr>
  <% if Froogle_Submit<>"" or not(isNull(Froogle_Submit)) then %>
  <tr bgcolor='#FFFFFF'>
  <td width="25%" class="inputname" height=20><B>Send Feed Tonight</B></td>
  <td width="75%" class="inputvalue">
		  <input class="image" type="checkbox" <%= force_checked %> name="Force_Resend">
		 <% small_help "force_resend" %></td>
  </td>
  </tr>
  <% end if %>

  <tr bgcolor='#FFFFFF'>
  <td height=20>
  <B>Google Base Feed File</b></td><td colspan=2>
  <% Set fso = CreateObject("Scripting.FileSystemObject")
  Froogle_Folder = fn_get_sites_folder(Store_Id,"Root")
  if fso.fileexists(Froogle_Folder&"/froogle.txt") then
     Set f = fso.getfile(Froogle_Folder&"/froogle.txt")
     sDate=f.DateLastModified
     set fso=Nothing
     set f=Nothing
  %>

		  <a href="<%= Site_Name %>froogle.txt" target=_blank class=link>Click here to download the file created for Froogle</a> : <B>File Created</b> <%= sDate %>
  <% else %>
  Your Google Base Feed File has not yet been created or uploaded.  Please note that this is done once every 24 hours normally between the hours of 12:00AM-4:00AM PST.  Please check back in 24 hours
  <% end if %>
  </td>
  </tr>
  
<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Froogle_Username","req","Please enter your froogle username.");
 frmvalidator.addValidation("Froogle_Password","req","Please enter your froogle password.");
 frmvalidator.addValidation("Froogle_Server","req","Please enter the froogle server address.");


</script>

