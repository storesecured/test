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
		sqlCustomFields="select * from store_forms where fldpbid=" & CustomField_ID & " and Store_Id=" &Store_Id
		rs_custom_fields=conn_store.execute(sqlCustomFields)

		if rs_custom_fields.EOF = false then
		  lclSubjectform=checkStringForQBack(rs_custom_fields("fldSubjectform"))
		  lclPageId = rs_custom_fields("fldpageid")
		  lclPagename=rs_custom_fields("fldPagename")
		  lclEmailTo=checkStringForQBack(rs_custom_fields("fldToEmail"))
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if
	set rs_custom_fields=nothing

	if op="edit" then
		sFullTitle = "Design > <a href=fields_page.asp class=white>Forms</a> > Edit - "&lclSubjectform
		sTitle = "Edit Form - "&lclSubjectform
	else
		sFullTitle = "Design > <a href=fields_page.asp class=white>Forms</a> > Add"
	        sTitle = "Add Form"
        end if
	
	sCommonName="Form"
	sCancel="fields_page.asp"
	sFormAction = "custom_fields_page_action.asp"
	thisRedirect = "custom_fields_page.asp"
'	thisRedirect = "item_add_page.asp"
	sFormName ="Create_Custom_Fields"
	sMenu = "design"
	addPicker=1
	sQuestion_Path = "advanced/custom_fields.htm"
	createHead thisRedirect
	
if Service_Type < 1 and Page_Id <> 5 then %>
	<tr bgcolor='#FFFFFF'>
		<td colspan=2>
			This feature is not available at your current level of service, you may only edit the home page.<BR><BR>
			PEARL Service or higher is required.
			<a href=billing.asp class=link>Click here to upgrade now.</a>
		</td>
	</tr>
	<% createFoot thisRedirect, 0%>
<% else %>
	<input type="hidden" name="op" value="<%=op%>">
	<input type="hidden" name="CustomField_ID" value="<%=CustomField_ID%>">
		 
			<%if op="edit" then%><tr bgcolor='#FFFFFF'>
				<td width="100%" colspan="3" height="11">

				<!--<input type="button" OnClick=JavaScript:self.location="/Item_Add_page.asp?id=<%=CustomField_ID%>" class="Buttons" value="Add Form  Fields" name="Create_new_Field">
				<input type="button" OnClick=JavaScript:self.location="/view_fields_items.asp?id=<%=CustomField_ID%>" class="Buttons" value="View Edit Fields" name="viewEdit_Field">-->
				<input type="button" OnClick=JavaScript:self.location="/form_fields_list.asp?fldPBID=<%=CustomField_ID%>" class="Buttons" value = "Form Field List" name="viewEdit_Field">


				</td>
			</tr>
                        <%End if %>
		
			<tr bgcolor='#FFFFFF'>
				<td width="25%" class="inputname"><B>Subject </B></td>
				<td width="75%" class="inputvalue">
					<INPUT  name="Subject" size="60" maxlength="100" value="<%=lclSubjectform%>">
					<% small_help "Subject" %>
				</TD>
				
			</TR>

			<tr bgcolor='#FFFFFF'>
				<td width="25%" class="inputname"><B>Page Name</B></td>
				
				<%if op="edit" then%>

					<td>
						<%=lclPagename%><input name="pageselect" type="hidden" value="<%=lclPageId%>"><% small_help "Page Name" %>
					</td>
				
				<%else%>

					<td width="75%" class="inputvalue">
						<% str_select = "SELECT page_id,Page_name FROM Store_pages WHERE Store_ID =" & store_id &" and is_link= 0 order by page_name"
						rs_store.open str_select,conn_store,1,1
						 %>
						<select name="pageselect">
						<%Do while not rs_store.eof%>
						<%if lclPagename =checkStringForQBack(rs_store("Page_name")) then %>
							<option value=<%=rs_store("page_id")%> selected ><%=checkStringForQBack(rs_store("Page_name"))%></option>	
						<%else%>	
							<option value=<%=rs_store("page_id")%>><%=checkStringForQBack(rs_store("Page_name"))%></option>
							
						<%end if
						rs_store.movenext
						  loop
						  rs_store.close%>
						</select>
						<% small_help "Page Name" %>
					</td>

				<%end if%>

			</tr>
			
			<tr bgcolor='#FFFFFF'>
				<td width="25%" class="inputname"><B>To Email </B></td>
				<td width="75%" class="inputvalue">
					<INPUT  name="Emailto" size="60" maxlength=100 value="<%=lclEmailTo%>">
					<% small_help "Email To" %>
				</TD>
			</TR>

	<% 
	'Response.End 
	createFoot thisRedirect, 1
	function checkStringForQBack(theString)
		tmpResult = theString
		tmpResult = replace(tmpResult,"&#8243;","&Prime;")
		tmpResult = replace(tmpResult,"&#8242;","&prime;")
		checkStringForQBack = tmpResult
	end function%>

<SCRIPT language="JavaScript">
	var frmvalidator  = new Validator(0);
	frmvalidator.addValidation("Subject","req","Please enter Subject.");
	frmvalidator.addValidation("Emailto","req","Please enter to Email");
	//frmvalidator.addValidation("View_Order","req","Please enter a view order.");
</script>
<%end if %>
