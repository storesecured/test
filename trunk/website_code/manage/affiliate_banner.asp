<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%
View_Order=0
op=Request.QueryString("op")
if op="edit" then
		Banner_Id = Request.QueryString("Id")
	if not isNumeric(Banner_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	
	sqlBanners="select * from Store_Affiliate_Banners  where Banner_Id=" & Banner_Id & " and Store_Id=" &Store_Id
	rs_banners=conn_store.execute(sqlBanners)

	
	if rs_banners.EOF = false then
'	  Banner_Image=checkStringForQBack(rs_banners("Banner_Image"))
	  Banner_Image=rs_banners("Banner_Image")
	  Banner_Text=rs_banners("Banner_Text")
	  View_Order=rs_banners("view_order")
	else
		Response.Redirect "admin_error.asp?message_id=1"
	end if

end if

if op="edit" then
	sTitle = "Marketing > Affiliates > Banners > Edit"
else
	sTitle = "Marketing > Affiliates > Banners > Add"
end if

sFormAction = "affiliate_banner_action.asp"
sFormName = "affiliate_banner"
thisRedirect = "affiliate_banner.asp"
addPicker = 1
sMenu = "marketing"
createHead thisRedirect
if service_type < 7	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>

		<input type="hidden" name="op" value="<%=op%>">
		<input type="hidden" name="Banner_Id" value="<%=Banner_Id%>">

		<tr bgcolor='#FFFFFF'><TD colspan=3>Use this page to add banners to show your affiliates.</td></tr>

			<tr bgcolor='#FFFFFF'>
					 <td width="100%" colspan="3" height="11">
						 <input type="button" OnClick=JavaScript:self.location="banners_list.asp" class="Buttons" value="Back to Affiliate Banners" name="Create_New_Banner"></td>
			</tr>
				 <% end if %>
			
			<tr bgcolor='#FFFFFF'>
				<td width="25%" height="25" class="inputname"><B>Banner Image</b></td>
				<td width="75%" height="25" class="inputvalue">
				<input type="text" name="Banner_Image" value="<%= Banner_Image %>" size="60">
				<a href="javascript:goImagePicker('Banner_Image')"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
				<a class="link" href="JavaScript:goFileUploader('Banner_Image');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Banner Image" %></td>
		</tr>

		<tr bgcolor='#FFFFFF'>
			<td width="25%" height="25" class="inputname"><B>Banner Text</b></td>
			<td width="75%" height="25" class="inputvalue">
				<input type="text" name="Banner_Text" value="<%= Banner_Text %>" size="60" maxlength="255">
				<INPUT type="hidden"  name=Banner_Text_C value="Op|String|0|255|||Banner Text">
				<% small_help "Banner Text" %></td>
		</tr>


		<tr bgcolor='#FFFFFF'>
			<td width="25%" height="25" class="inputname"><B>View Order</b></td>
			<td width="75%" height="25" class="inputvalue">
				<input type="text" name="View_Order" value="<%= View_Order %>" size="3" maxlength=3 onKeyPress="return goodchars(event,'0123456789')">
				<INPUT type="hidden"  name=View_Order_C value="Op|String|1|3|||View Order">
				<% small_help "View Order" %></td>
		</tr>

		
		<input type="hidden" name="Banner_Id" value="<%=Banner_Id%>">

		
<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
frmvalidator.addValidation("Banner_Text","req","Please enter Text.");
frmvalidator.addValidation("View_Order","req","Please enter a View Order.");
 </script>

