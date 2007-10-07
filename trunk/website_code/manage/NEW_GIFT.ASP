<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<% 

gift_name = ""
gift_amount = ""
gift_validity = "30"
gift_item = ""
gift_restricted = ""

if request.querystring("ID")<>"" or request.querystring("gift_id") <> "" then
	Gift_ID = request.querystring("ID")
	if Gift_ID = "" then
	Gift_ID = request.querystring("Gift_ID")
	end if
	if not isNumeric(Gift_ID) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sql_select = "select * from store_gift_certificates where gift_id = "&Gift_ID&" and store_id = "&store_id
	rs_Store.open sql_select, conn_store, 1, 1 

	if rs_Store.bof and rs_Store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	else
	gift_name = rs_Store("Gift_Name")
	gift_amount = rs_Store("Amount")
	gift_validity = rs_Store("Validity_Time")
	gift_item = rs_Store("Item")
	gift_restricted = rs_Store("Restricted_Items")
	end if
	rs_Store.close

End If

if Gift_ID<>"" then
   sTitle = "Edit Gift Certificate - "&gift_name
   sFullTitle = "Marketing > <a href=gift_manager.asp class=white>Gift Certificates</a> > Edit - "&gift_name
   sFormName ="Update_Gift"
else
   sTitle = "Add Gift Certificate"
   sFullTitle = "Marketing > <a href=gift_manager.asp class=white>Gift Certificates</a> > Add"
   sFormName="Create_Gift"
end if

sCommonName="Gift Certificate"
sCancel="gift_manager.asp"
addPicker=1
sFormAction = "Gift_action.asp"
sName= "new_gift"
thisRedirect = "new_gift.asp"
sMenu = "marketing"
sQuestion_Path = "marketing/gift_certificates.htm"
createHead thisRedirect
%>




<% if Gift_ID<>"" then %>
	<input type="hidden" name="Gift_ID" value="<%= Gift_ID %>">
<% End If %>


	<tr bgcolor='#FFFFFF'>
	<td width="100%" colspan="3" height="11">
			<input type="button" OnClick=JavaScript:self.location="Coupon_settings.asp" class="Buttons" value="Coupon Settings" name="Create_new_Coupon">

    
	</tr>
    <% if Show_Coupon =0 then %>
			<tr bgcolor='red'>
				    <td width="100%" colspan="4" height="11">
					    <font class=white>Coupons are currently disabled in your store.  Your customers will be unable to redeem this gift certificate or coupons unless coupons are enabled.  Go to the coupon settings page to enable coupons.</font></td>
			</tr>
			<% end if %>


	    <tr bgcolor='#FFFFFF'>
		    <td width="10%" class="inputname"><b>Name</b></td>
		    <td width="90%" class="inputvalue">
			    <input type="text" name="Gift_name" size="60" maxlength=50 value="<%= gift_name %>">
			    <INPUT type="hidden"  name=Gift_name_C value="Re|String|0|50|||Name">
			    <% small_help "Name" %></td>
    </tr>

	    <tr bgcolor='#FFFFFF'>
		    <td width="10%" class="inputname"><b>Amount</b></td>
		    <td width="90%" class="inputvalue">
			    <%= Store_Currency %><input type="text" name="Gift_Amount" size="10" onKeyPress="return goodchars(event,'0123456789.')" value="<%= gift_amount %>">
			    <INPUT type="hidden"  name=Gift_Amount_C value="Re|Integer|0|32767|||Amount">
			    <% small_help "Amount" %></td>
    </tr>

	    <tr bgcolor='#FFFFFF'>
		    <td width="10%" class="inputname"><b>Validity</b></td>
		    <td width="90%" class="inputvalue">
			    <input type="text" name="Gift_Validity" size="10" onKeyPress="return goodchars(event,'0123456789')" value="<%= gift_validity %>">
			    <INPUT type="hidden"  name=Gift_Validity_C value="Re|Integer|0|32767|||Validity"> days
			    <% small_help "Validity" %></td>
    </tr>

	    <tr bgcolor='#FFFFFF'>
		    <td width="10%" class="inputname"><b>Gift Certificate Item</b></td>
		    <td width="90%" class="inputvalue">
			    <input type="text" name="Gift_Item" size="10" value="<%= gift_item %>" onKeyPress="return goodchars(event,'0123456789')">
			    <a class="link" href="JavaScript:goItemPicker('Gift_Item');">Item Picker</a>
			    <input type="hidden" name="Gift_Item_Value" value="<%= gift_item %>">
			    <INPUT type="hidden"  name=Gift_Item_C value="Re|Integer|||||Gift Certificate Item">
			    <BR><small>The gift certificate item is the item in your inventory
             which when purchased should generate a gift certificate code.  Ie if this is a $25 gift certificate you would make an inventory item
             called $25 gift certificate with a price of $25 and enter the item above.  Now when that item is purchased it will automatically make a gift certificate for the above amount and send the code to the customer.</small><% small_help "Gift Certificate Item" %></td>
    </tr>

	    <tr bgcolor='#FFFFFF'>
		    <td width="10%" class="inputname"><b>Restricted Items</b></td>
		    <td width="90%" class="inputvalue">
			    <input type="text" name="Gift_Restricted" size="10" maxlength=50 value="<%= gift_restricted %>" onKeyPress="return goodchars(event,'0123456789,')">
			    <a class="link" href="JavaScript:goItemPicker('Gift_Restricted');">Item Picker</a>
			    <INPUT type="hidden"  name=Gift_Restricted_C value="Op|String|0|50|||Restricted Items">
			    <% small_help "Restricted Items" %></td>
    </tr>

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Gift_name","req","Please enter a name for this gift certificate.");
 frmvalidator.addValidation("Gift_Amount","req","Please enter a amount.");
 frmvalidator.addValidation("Gift_Item","req","Please enter the item id.");
 frmvalidator.addValidation("Gift_Validity","req","Please enter the number of days the gift certificate is valid.");

</script>
