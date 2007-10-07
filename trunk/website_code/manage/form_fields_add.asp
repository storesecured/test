<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
sFormAction = "form_fields_action.asp"
sFormName = "form_fields_add.asp"
CustomField_ID = Request.QueryString("ID")
sFullTitle = "Design > <a href=fields_page.asp class=white>Forms</a> > <a href=form_fields_list.asp?fldpbid="&CustomField_ID&" class=white>Fields</a> > "

op=Request.QueryString("op")








	if not isNumeric(CustomField_ID) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if


if op="edit" then
                sqlCustomFields="select * from store_form_fields where fldid=" & CustomField_ID & " and Store_Id=" &Store_Id
		rs_custom_fields=conn_store.execute(sqlCustomFields)

		if rs_custom_fields.EOF = false then
		'  lclSubjectform=checkStringForQBack(rs_custom_fields("fldSubjectform"))
		 ' lclPageId = rs_custom_fields("fldpageid")
		  Custom_Field_Name=rs_custom_fields("fldname")
		  Custom_Field_Type=rs_custom_fields("fldType")
		  
		  Custom_Field_Values=rs_custom_fields("fldFieldType")
		  Field_Required=rs_custom_fields("fldRequired")
		  View_Order=rs_custom_fields("fldViewOrder")
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		
		sFullTitle=sFullTitle&"Edit - "&Custom_Field_Name
		sTitle = "Edit Form Field - "&Custom_Field_Name

else
     sTitle="Add Form Field"
        sFullTitle = sFullTitle & "Add"
end if

sCancel="form_fields_list.asp?fldpbid="&CustomField_ID
sCommonName="Form Field"
thisRedirect = "form_fields_add.asp"
sMenu="design"
createHead thisRedirect
'=================ddddddddd
%>
	<tr bgcolor='#FFFFFF'>
		<td width="25%" class="inputname"><B>Name</b></td>
		<td width="75%" colspan="2" class="inputvalue">
			<input type="text" name="Custom_Field_Name" size="60" maxlength=100 value="<%= Custom_Field_Name %>">
			<INPUT type="hidden"  name=Custom_Field_Name_C value="Re|String|0|100|||Name">
			<% small_help "Name" %>
		</td>
	</tr>
	<tr bgcolor='#FFFFFF'>
		<td width="10%" class="inputname" rowspan=4><B>Type</B></td>
		<td width="25%" class="inputvalue">
			<input class="image" type="radio" value="1" name="Custom_Field_Type"
			<% if op="edit" then %>
				<% if Custom_Field_Values = 1 then %>
					checked
				<% end if %>
			<% else %>
				checked
			<% end if %>
			>Text Field
		</td>
		<td width="65%" class="inputvalue">
			size <input type="text" name="Custom_Text_Value"
			<% if op="edit" then %>
				<% if Custom_Field_Values = 1  then %>
					value="<%= Custom_Field_Type %>"
				<% end if %>
			<% end if %>			    
			size="3" onKeyPress="return goodchars(event,'0123456789.')">
			<%small_help "Text Field" %>
		</td>
	</tr>
			
	<tr bgcolor='#FFFFFF'>
		<td width="25%" class="inputvalue">
			<input class="image" type="radio" value="2" name="Custom_Field_Type"
			<%if op="edit" then%>
				<% if Custom_Field_Type=2 then %>
					checked
				<% end if %>
			<% end if %>				  
			>Text Area
		</td>
		<td width="65%" class="inputvalue">
			cols<input type="text" name="Custom_Text_Area_Value"
			<%if op="edit" then %>
				<% if Custom_Field_Values=2 then %>
				value="<%= Custom_Field_Type %>"
				<%else %>
				value="0"
				<%end if%>
			<%else %>
				value="0"
			<%end if %>				  
			size="3" maxlength="5" onKeyPress="return goodchars(event,'0123456789.')">
            <%small_help "Text Area" %>
         </td>
	</tr>
	
	<tr bgcolor='#FFFFFF'>
		<td width="25%" class="inputvalue">
			<input class="image" type="radio" value="3" name="Custom_Field_Type"
			<%if op="edit" then %>
			
				<%if Custom_Field_Values=3 then %>
				checked
				<%end if %>
			<%end if %>				  
			>Combo Box
		</td>
				
		<td width="65%" class="inputvalue">
			<input type="text" name="Custom_Combo_Value"
			<%if op="edit" then %>
				<%if Custom_Field_Values=3 then %>
				value="<%= Custom_Field_Type %>"
				<% end if %>
			<% end if %>
			size="60" maxlength="3000">
			<INPUT type="hidden"  name=Custom_Combo_Value_C value="Op|String|0|4000|||Combo Box">
	       <%small_help "Combo Box" %>
	     </td>
	</tr>


	<tr bgcolor='#FFFFFF'>
		<td width="25%" class="inputvalue">
			<input class="image" type="radio" value="4" name="Custom_Field_Type"
			<%if op="edit" then %>
			
				<%if Custom_Field_Values=4then %>
				checked
				<%end if %>
			<%end if %>				  
			>Check Box
		</td>
				
		<td width="65%" class="inputvalue">
	       <%small_help "Check Box" %>
	     </td>
	</tr>	

	<input type="hidden" name="CustomField_ID" value="<%=CustomField_ID%>">

	<tr bgcolor='#FFFFFF'>
		<td width="25%" class="inputname"><B>Field Required</B></td>
		<td width="75%" class="inputvalue" colspan="2">
			<input class="image" type="checkbox" name="Field_Required"
			<% if Field_Required then %>
				checked
			<% end if %>>
			<% small_help "Field Required" %>
		</td>
	</tr>
			
	 <tr bgcolor='#FFFFFF'>
		<td width="25%" class="inputname"><B>View Order</B></td>
		<td width="75%" class="inputvalue" colspan="2" >
			<input name="View_Order" size="3" value="<%= View_Order %>" onKeyPress="return goodchars(event,'0123456789.')">
			<INPUT type="hidden"  name="View_Order_C" value="Re|Integer|||||View Order">
			
			<% small_help "View Order"%>
		</td>
		<input type="hidden" name="op" value="<%=op%>">		
	</tr>
  <% createFoot thisRedirect, 1%>
  
<SCRIPT language="JavaScript">
	var frmvalidator  = new Validator(0);
	frmvalidator.addValidation("Custom_Field_Name","req","Please enter a field name.");
	frmvalidator.addValidation("Custom_Field_Type","req","Please enter a field type.");
	//frmvalidator.addValidation("Custom_Field_Values","req","Please enter a field value.");
	frmvalidator.addValidation("View_Order","req","Please enter a view order.");
</script>
