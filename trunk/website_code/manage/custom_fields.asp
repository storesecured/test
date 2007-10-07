<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
View_Order=0
op=Request.QueryString("op")
if op="edit" then
		CustomField_ID = Request.QueryString("ID")
	if not isNumeric(CustomField_ID) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	
	sqlCustomFields="select * from Store_Custom_Fields  where CustomField_ID=" & CustomField_ID & " and Store_Id=" &Store_Id
	rs_custom_fields=conn_store.execute(sqlCustomFields)

	
	if rs_custom_fields.EOF = false then
	  Custom_Field_Name=checkStringForQBack(rs_custom_fields("custom_field_name"))
	  Custom_Field_Type=rs_custom_fields("custom_field_type")
	  Custom_Field_Values=rs_custom_fields("custom_field_values")
	  Field_Required=rs_custom_fields("field_required")
	  View_Order=rs_custom_fields("view_order")
	else
		Response.Redirect "admin_error.asp?message_id=1"
	end if

end if


set rs_custom_fields=nothing

sTextHelp="fields/custom_fields.doc"
sCommonName="Custom Field"
sCancel="fields.asp"
if op="edit" then
	sTitle = "Edit Custom Field - "&Custom_Field_Name
	sFullTitle = "General > <a href=fields.asp class=white>Custom Fields</a> > Edit - "&Custom_Field_Name
else
	sTitle = "Add Custom Field"
	sFullTitle = "General > <a href=fields.asp class=white>Custom Fields</a> > Add"
end if
sFormAction = "custom_fields_action.asp"
thisRedirect = "custom_fields.asp"
sFormName ="Create_Custom_Fields"
sMenu = "general"
addPicker=1
sQuestion_Path = "advanced/custom_fields.htm"
createHead thisRedirect
if Service_Type < 5 then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service, you may only edit the home page.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>



<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="CustomField_ID" value="<%=CustomField_ID%>">


			<TR bgcolor='#FFFFFF'>
				<td width="25%" class="inputname"><B>Name</b></td>
				<td width="75%" colspan="2" class="inputvalue">
					<input type="text" name="Custom_Field_Name" size="60" maxlength=200 value="<%= Custom_Field_Name %>" onKeyPress="return goodchars(event,'-0123456789.abcdefghijklmnopqrstuvwxyz_ ?<>+,$#@!%^&*()=|}{[]/~:;')">
						<INPUT type="hidden"  name=Custom_Field_Name_C value="Re|String|0|200||<,>|Name">
					<% small_help "Name" %></td>
			</tr>




			<TR bgcolor='#FFFFFF'>
				<td width="10%" class="inputname" rowspan=4><B>Type</B></td>
				<td width="25%" class="inputvalue">
						<input class="image" type="radio" value="1" name="Custom_Field_Type"
							<% if op="edit" then %>
								<% if Custom_Field_Type = 1 then %>
									checked
								<% end if %>
							<% else %>
								checked
							<% end if %>
						>Text Field</td>
				<td width="65%" class="inputvalue">
						<input type="text" name="Custom_Text_Value"
							<% if op="edit" then %>
								<% if Custom_Field_Type = 1  then %>
									value="<%= Custom_Field_Values %>"
								<% end if %>
							<% end if %>			    
						size="3" onKeyPress="return goodchars(event,'0123456789.')">
						<!--<INPUT type="hidden" name=Custom_Text_Value_C value="Re|Integer|||||Text Field">-->
                              <% small_help "Text Field" %></td>
			</tr>
    
				<TR bgcolor='#FFFFFF'>
				<td width="25%" class="inputvalue">
						<input class="image" type="radio" value="2" name="Custom_Field_Type"
							<% if op="edit" then %>
								<% if Custom_Field_Type=2 then %>
									checked
								<% end if %>
							<% end if %>				  
						>Text Area</td>
				<td width="65%" class="inputvalue">
						<input type="text" name="Custom_Text_Area_Value"
							<% if op="edit" then %>
								<% if Custom_Field_Type=2 then %>
									value="<%= Custom_Field_Values %>"
									<% else %>
									value="0"
								<% end if %>
								<% else %>
									value="0"
							<% end if %>				  
						size="3" maxlength="5" onKeyPress="return goodchars(event,'0123456789.')">
<!--						<INPUT type="hidden"  name=Custom_Text_Area_Value_C value="Re|Integer|||||Text Area">-->
                              <% small_help "Text Area" %></td>
			</tr>



				<TR bgcolor='#FFFFFF'>
				<td width="25%" class="inputvalue">
						<input class="image" type="radio" value="3" name="Custom_Field_Type"
							<% if op="edit" then %>
								<% if Custom_Field_Type=3 then %>
									checked
								<% end if %>
							<% end if %>				  
						>Combo Box</td>
				<td width="65%" class="inputvalue">
						<input type="text" name="Custom_Combo_Value"
							<% if op="edit" then %>
								<% if Custom_Field_Type=3 then %>
									value="<%= Custom_Field_Values %>"
								<% end if %>
							<% end if %>				  
						size="30" maxlength="1000">
	<!--					<INPUT type="hidden"  name=Custom_Combo_Value_C value="Re|String|0|1000|||Combo Box">-->
                              <% small_help "Combo Box" %></td>
			</tr>
			
			<TR bgcolor='#FFFFFF'>
				<td width="25%" class="inputvalue">
						<input class="image" type="radio" value="4" name="Custom_Field_Type"
							<% if op="edit" then %>
								<% if Custom_Field_Type=4 then %>
									checked
								<% end if %>
							<% end if %>				  
						>Checkbox</td>
				<td width="65%" class="inputvalue">&nbsp;
						<% small_help "CheckBox" %></td>
			</tr>



				<TR bgcolor='#FFFFFF'>
				<td width="25%" class="inputname"><B>Field Required</B></td>
				<td width="75%" class="inputvalue" colspan="2" >
						<input class="image" type="checkbox" name="Field_Required"
							<% if Field_Required then %>
								checked
							<% end if %>>
				 <% small_help "Field Required" %></td></tr>
				 
				 <TR bgcolor='#FFFFFF'>
					<td width="25%" class="inputname"><B>View Order</B></td>
					<td width="75%" class="inputvalue" colspan="2" >
						<input name="View_Order" size="3" value="<%= View_Order %>"  onKeyPress="return goodchars(event,'0123456789.')">
							<INPUT type="hidden"  name="View_Order_C" value="Re|Integer|||||View Order">
							<% small_help "View Order" %></td>
				</tr>


		
				
<% createFoot thisRedirect, 1

function checkStringForQBack(theString)
	tmpResult = theString
	tmpResult = replace(tmpResult,"&#8243;","&Prime;")
	tmpResult = replace(tmpResult,"&#8242;","&prime;")
	checkStringForQBack = tmpResult
end function
%>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

frmvalidator.addValidation("Custom_Field_Name","req","Please enter a field name.");
//frmvalidator.addValidation("Custom_Field_Type","req","Please enter a field type.");
//frmvalidator.addValidation("Custom_Field_Values","req","Please enter a field value.");
frmvalidator.addValidation("View_Order","req","Please enter a view order.");


</script>
<% end if %>
