<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/fonts.asp"-->
<script language="JavaScript" src="include/picker.js"></script>

<script language="JavaScript">
	function ChangeColor1(src, obj)
	{
		var hxv =document.getElementById(src).value;
		if (hxv.length == 6)
		{
				document.getElementById(obj).style.backgroundColor = hxv;
		}
	}
	</script>
<%


Template_Id=Request.Querystring("Id")
If Template_Id = "" then
	fn_error "Please select a template"
else

		sql_check="select template_name, template_active from store_design_template where store_id="&Store_Id&" and template_id="&Template_Id
		set rs_check=conn_store.execute(sql_check)

		'EOF check
		if rs_check.EOF then
		   fn_error "This template does not exist"
		else

				templ_active_check=rs_check("template_active")
				templ_name=rs_check("template_name")
				set rs_check=nothing

				sql_select = "select Global_Bck_Color, Global_Text_Font_Color, Global_Text_Font_Face, Global_Text_Font_Size , Global_Link_Font_Color, Global_Link_Font_Face, Global_Link_Font_Size, Global_Nav_Font_Color, Global_Nav_Font_Face, Global_Nav_Font_Size from store_design_template where store_id="&Store_Id&" and template_id="&Template_Id

				set rs_select = conn_store.execute(sql_select)
				if not rs_select.EOF then
					Global_Bck_Color = rs_select("Global_Bck_Color")
					Global_Text_Font_Color = rs_select("Global_Text_Font_Color")
					Global_Text_Font_Face = rs_select("Global_Text_Font_Face")
					Global_Text_Font_Size = rs_select("Global_Text_Font_Size")
					Global_Link_Font_Color = rs_select("Global_Link_Font_Color")
					Global_Link_Font_Face = rs_select("Global_Link_Font_Face")
					Global_Link_Font_Size = rs_select("Global_Link_Font_Size")
					Global_Nav_Font_Color = rs_select("Global_Nav_Font_Color")
					Global_Nav_Font_Face = rs_select("Global_Nav_Font_Face")
					Global_Nav_Font_Size = rs_select("Global_Nav_Font_Size")
				end if

				sFormAction = "process_form.asp"
				sName = "EditImage"
				sFormName = "bck_font"
				sTitle = "Edit Template Color and Text - "&templ_name
				sFullTitle = "Design > <a href=template_list.asp class=white>Template</a> > <a href=layout_design.asp?Id="&Template_Id&" class=white>Edit</a> > Color and Text - "&templ_name
				sCommonName = "Color and Text"
		                sCancel = "layout_design.asp?Id="&Template_Id
                                sSubmitName = "Update"
				thisRedirect = "bck_font.asp"
				addPicker = 1
				sMenu = "design"
				sQuestion_Path = "design/color_and_text.htm"
				createHead thisRedirect
				if Service_Type < 1	then %>
					<TR bgcolor='#FFFFFF'>
					<td colspan=2>
						This feature is not available at your current level of service.<BR><BR>
						PEARL Service or higher is required.
						<a href=billing.asp class=link>Click here to upgrade now.</a>
					</td></tr>
					<% createFoot thisRedirect, 0%>

				<% else %>

					<TR bgcolor='#FFFFFF'>
						 <td width="100%" colspan="3" height="11">
						 <input type="button" OnClick=JavaScript:self.location="template_list.asp" class="Buttons" value="Template List" name="Create_new_Page">&nbsp;<input type="button" OnClick=JavaScript:self.location="layout_design.asp?op=edit&Id=<%=Template_ID%>" class="Buttons" value="Template Detail" name="Create_new_Page"></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
							<td colspan="3" align="center">
								<BR>
								<font color="#ff0000" face="verdana" size="1"><b>Click the Apply Template button to apply this design to your store.</b></font>

							</td>
						</tr>
				<!--Added for bringing the Preview and apply Template buttons-->
				<TR bgcolor='#FFFFFF'>
					<td colspan=3 align="center">
							<table cellspacing="5" cellpadding="0">
									<TR bgcolor='#FFFFFF'>
										<td>	
											<!--<a href="layout_settings.asp?Gen=1&template_id=<%=Template_Id%>">Apply Template</a>-->
											<input type="button" class="buttons" value="Apply Template" onClick="window.location='layout_settings.asp?Gen=1&template_id='+<%=Template_Id%>">							
										</td>
										<td>
											<!--<a href="layout_settings.asp?Gen=1&preview_id=1&template_id=<%=Template_Id%>">Preview Template</a>-->
											<input type="button" class="buttons" value="Preview Template" onClick="window.location='layout_settings.asp?Gen=1&preview_id=1&template_id='+<%=Template_Id%>">
										</td>
									</tr>
								</table>
					</td>
				</TR>
				<!--Added for bringing the Preview and apply Template buttons-->

<!--						<TR bgcolor='#FFFFFF'>
							<td colspan="3" align="center">
								<table cellspacing="5" cellpadding="0">
									<TR bgcolor='#FFFFFF'>
										<td>	
											<form name="frm_gendesign1" method="post" action="layout_settings.asp">
												<input type="hidden" name="Gen" value="1">
												<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
												<%If templ_active_check=0 then%>
													<input class="buttons" type="button" name="GenSubmit" value="Apply Template Now!" OnClick='if (confirm("Are you sure you want to apply this template to your store?")) JavaScript:document.frm_gendesign1.submit();' >
												<% else %>
													<input class="buttons" type="button" name="GenSubmit" value="Apply Template" OnClick='if (confirm("Are you sure you want to apply this template to your store?")) JavaScript:document.frm_gendesign1.submit();' >
												<% end if %> 
											</form>
										</td>
										<td>
											<form name="frm_gendesign2" method="post" action="layout_settings.asp">
												<input type="hidden" name="Gen" value="1">
												<input type="hidden" name="Preview_Id" id="Preview_Id" value="1">
												<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
												<input class="buttons" type="submit" name="GenSubmit" value="Preview Template">
											</form>
										</td>
									</tr>
								</table>
							</td>
						</tr>
-->
					<TR bgcolor='#FFFFFF'><TD colspan=3>Use this page to modify the text properties and/or background colors or images.</td></tr>

					 <TR bgcolor='#FFFFFF'>
						<td width="7%" height="25" class="inputname"><B>Background</b><BR>color or file</td>
								<td width="93%" height="25" class="inputvalue">
							<input type="text" name="Global_Bck_Color" value="<%= Global_Bck_Color %>" size="20">
							<INPUT type="hidden"  name=Global_Bck_Color_C value="Op|String|0|50|||Background">
							<a href="javascript:goColorPicker('Global_Bck_Color')"><img border="0" src="images/bgcol.gif" width="23" height="22" alt="Color Picker"></a>
							<a href="javascript:goImagePicker('Global_Bck_Color')"><img border="0" src="images/image.gif" width="23" height="22" alt="Image Picker"></a>
							<a class="link" href="JavaScript:goFileUploader('Global_Bck_Color');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>

							<% small_help "Background" %></td>


					</tr>

					<TR bgcolor='#FFFFFF'>
			<td width="7%" height="25" class="inputname"><B>Text Color</b></td>
			<td nowrap width="50%" height="25" class="inputvalue">
				<input type="text" size=7 maxlength=7 name="Global_Text_Font_Color" id="Global_Text_Font_Color" value="<%= Global_Text_Font_Color %>" size="20" maxlength=6 onKeyPress="return goodchars(event,'0123456789ABCDEF')" onKeyUp="ChangeColor1('Global_Text_Font_Color','colorPreview');" onFocus="ChangeColor1('Global_Text_Font_Color','colorPreview'); "onBlur="ChangeColor1('Global_Text_Font_Color','colorPreview');">
				<!--Preview Box-->
				<%
					if len(Global_Text_Font_Color) <> 0 then
					%>
						<input type="text"  size=1 id="colorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= Global_Text_Font_Color %>" >		
					<%
					end if
				%>
				
				<!--Preiview Box END -->
				<INPUT type="hidden"  name=Global_Text_Font_Color_C value="Op|String|6|7|||Text Color">
				<a href="javascript:goColorPicker('Global_Text_Font_Color')"><img border="0" src="images/fontcolor.gif" width="23" height="22" alt="Color Picker"></a>
				<% small_help "Text Color" %></td>
		</tr>

					<TR bgcolor='#FFFFFF'>
						<td width="7%" height="25" class="inputname"><B>Text Font</b></td>
						<td nowrap width="50%" height="25" class="inputvalue">
							<% font_selector Global_Text_Font_Face,"Global_Text_Font_Face",Short%>
							<% small_help "Text Font" %></td>
					</tr>

					<TR bgcolor='#FFFFFF'>
							<td width="7%" height="25" class="inputname"><B>Text Size</b><BR>1-5</td>
						<td width="93%" height="25" class="inputvalue">
							<input type="text" name="Global_Text_Font_Size" value="<%= Global_Text_Font_Size %>" size="20" onKeyPress="return goodchars(event,'0123456789.')">
							<INPUT type="hidden"  name=Global_Text_Font_Size_C value="Op|Integer|1|5|||Text Size">
							<% small_help "Text Size" %></td>
					</tr>

					<TR bgcolor='#FFFFFF'>
			<td width="7%" height="25" class="inputname"><B>Link Color</b></td>
			<td nowrap width="50%" height="25" class="inputvalue">
				
				<input type="text" size=7 maxlength=7 name="Global_Link_Font_Color" id="Global_Link_Font_Color"     value="<%= Global_Link_Font_Color %>" size="20" maxlength=6 onKeyPress="return goodchars(event,'0123456789ABCDEF')" onKeyUp="ChangeColor1('Global_Link_Font_Color','LinkColorPreview');" onFocus="ChangeColor1('Global_Link_Font_Color','LinkColorPreview'); " onBlur="ChangeColor1('Global_Link_Font_Color','LinkColorPreview');">

				<% 'COLOR PREVIEW BOX
				if len(Global_Link_Font_Color) <> 0 then
				%>
					<input type="text"  size=1 id="LinkColorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= Global_Link_Font_Color %>" >		
						
				<%
				end if
				'COLOR PREVIEW BOX END%>
				
				<INPUT type="hidden"  name=Global_Link_Font_Color_C value="Op|String|6|7|||Link Color">
				<a href="javascript:goColorPicker('Global_Link_Font_Color')"><img border="0" src="images/fontcolor.gif" width="23" height="22" alt="Color Picker"></a>
				

				
				<% small_help "Link Color" %></td>
		</tr>

					<TR bgcolor='#FFFFFF'>
						<td width="7%" height="25" class="inputname"><B>Link Font</b></td>
						<td nowrap width="50%" height="25" class="inputvalue">
							<% font_selector Global_Link_Font_Face,"Global_Link_Font_Face",Short%>
						<% small_help "Link Font" %></td></tr>

					<TR bgcolor='#FFFFFF'>
						<td width="7%" height="25" class="inputname"><B>Link Size</b><BR>1-5</td>
						<td width="93%" height="25" class="inputvalue">
							<input type="text" name="Global_Link_Font_Size" value="<%= Global_Link_Font_Size %>" size="20" onKeyPress="return goodchars(event,'0123456789.')">
							<INPUT type="hidden"  name=Global_Link_Font_Size_C value="Op|Integer|1|5|||Link Size">
							<% small_help "Link Size" %></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
			<td width="7%" height="25" class="inputname"><B>Nav Button Color</b></td>
			<td nowrap width="50%" height="25" class="inputvalue">
			
				<input type="text" size=7 maxlength=7 name="Global_Nav_Font_Color" id="Global_Nav_Font_Color"  value="<%= Global_Nav_Font_Color %>" size="20" maxlength=6 onKeyPress="return goodchars(event,'0123456789ABCDEF')" onKeyUp="ChangeColor1('Global_Nav_Font_Color','NavColorPreview');" onFocus="ChangeColor1('Global_Nav_Font_Color','NavColorPreview'); " onBlur="ChangeColor1('Global_Nav_Font_Color','NavColorPreview');">

				<% 'COLOR PREVIEW BOX
				if len(Global_Nav_Font_Color) <> 0 then
				%>
					<input type="text"  size=1 id="NavColorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= Global_Nav_Font_Color %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>
				
				<INPUT type="hidden"  name=Global_Nav_Font_Color_C value="Op|String|6|7|||Nav Color">
				<a href="javascript:goColorPicker('Global_Nav_Font_Color')"><img border="0" src="images/fontcolor.gif" width="23" height="22" alt="Color Picker"></a>
				<% small_help "Nav Color" %></td>
		</tr>

					<TR bgcolor='#FFFFFF'>
						<td width="7%" height="25" class="inputname"><B>Nav Button Font</b></td>
						<td nowrap width="50%" height="25" class="inputvalue">
							<% font_selector Global_Nav_Font_Face,"Global_Nav_Font_Face",Short%>
						<% small_help "Nav Font" %></td></tr>

					<TR bgcolor='#FFFFFF'>
						<td width="7%" height="25" class="inputname"><B>Nav Button Size</b><BR>1-5</td>
						<td width="93%" height="25" class="inputvalue">
							<input type="text" name="Global_Nav_Font_Size" value="<%= Global_Nav_Font_Size %>" size="20" onKeyPress="return goodchars(event,'0123456789.')">
							<INPUT type="hidden"  name=Global_Nav_Font_Size_C value="Op|Integer|1|5|||Nav Size">
							<% small_help "Nav Size" %></td>
					</tr>
					
						 <input type="hidden" name="Template_Id" value="<%= Template_Id%>">


						<% createFoot thisRedirect,1 %>
						<SCRIPT language="JavaScript">
						 var frmvalidator  = new Validator(0);
						 frmvalidator.addValidation("Global_Nav_Font_Size","lessthan=5.01","Please enter a nav button font size between 1 and 5");
						 frmvalidator.addValidation("Global_Nav_Font_Size","greaterthan=.99","Please enter a nav button font size between 1 and 5");
						 frmvalidator.addValidation("Global_Nav_Font_Color","minlength=6","Please enter a valid nav button color.");
						 frmvalidator.addValidation("Global_Link_Font_Size","lessthan=5.01","Please enter a link font size between 1 and 5");
						 frmvalidator.addValidation("Global_Link_Font_Size","greaterthan=.99","Please enter a link font size between 1 and 5.");
						 frmvalidator.addValidation("Global_Link_Font_Color","minlength=6","Please enter a valid link font color.");
						 frmvalidator.addValidation("Global_Text_Font_Size","lessthan=5.01","Please enter a return duration.");
						 frmvalidator.addValidation("Global_Link_Font_Size","greaterthan=.99","Please enter a return duration.");
						 frmvalidator.addValidation("Global_Text_Font_Color","minlength=6","Please enter a valid text color.");
						 </script>
					<% end if %>	
		<%end if	 ' EOF condition%>
<% end if		'Template_Id querystring condition%>
