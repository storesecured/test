<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/manage_object.asp"-->
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


if request.querystring("Del") = 1 then
	Object_Id=Request.QueryString("Id")
	Template_Id=Request.QueryString("Tid")
	sql_select="delete from store_design_objects where Object_Id="&Object_Id&" and Store_Id="&Store_Id&" and template_id="&Template_Id

	set rs_select=conn_store.execute(sql_select)
	response.redirect "layout_design.asp?Id="&Template_Id
else

        Template_Id=Request.QueryString("tid")
	sFullTitleStart = "Design > <a href=template_list.asp class=white>Template</a> > <a href=layout_design.asp?Id="&Template_Id&" class=white>Edit</a> > "
        sCommonName = "Object"
        sCancel = "layout_design.asp?Id="&Template_Id
		
	if Request.QueryString("ar") <> "" then

		ar=Request.QueryString("ar")

		select case ar
			case 1
				design_area_name="Top"
				design_area_BG="Design_Top_BG"
				design_area_border="Design_Top_border"
				design_area_height="Design_Top_height"
				design_area_width="Design_Top_width"
				design_area_border_color="Design_Top_Border_Color"
			case 2
				design_area_name="Left"
				design_area_BG="Design_Left_BG"
				design_area_border="Design_Left_border"
				design_area_height="Design_Left_height"
				design_area_width="Design_Left_width"
				design_area_border_color="Design_Left_Border_Color"
			case 3
				design_area_name="Right"
				design_area_BG="Design_Right_BG"
				design_area_border="Design_Right_border"
				design_area_height="Design_Right_height"
				design_area_width="Design_Right_width"
				design_area_border_color="Design_Right_Border_Color"
			case 4
				design_area_name="Bottom"
				design_area_BG="Design_Bottom_BG"
				design_area_border="Design_Bottom_border"
				design_area_height="Design_Bottom_height"
				design_area_width="Design_Bottom_width"
				design_area_border_color="Design_Bottom_Border_Color"
			case 5
				design_area_name="Center Top"
				design_area_BG="Design_Cen_Top_BG"
				design_area_border="Design_Cen_Top_border"
				design_area_height="Design_Cen_Top_height"
				design_area_width="Design_Cen_Top_width"
				design_area_border_color="Design_Cen_Top_Border_Color"
			case 6
				design_area_name="Center Bottom"
				design_area_BG="Design_Cen_Bot_BG"
				design_area_border="Design_Cen_Bot_border"
				design_area_height="Design_Cen_Bot_height"
				design_area_width="Design_Cen_Bot_width"
				design_area_border_color="Design_Cen_Bot_Border_Color"
			case 7
				design_area_name="Center Content"
				design_area_BG="Design_Cen_Cen_BG"
				design_area_border="Design_Cen_Cen_border"
				design_area_height="Design_Cen_Cen_height"
				design_area_width="Design_Cen_Cen_width"
				design_area_border_color="Design_Cen_Cen_Border_Color"
		end select

	
		sql_select="select template_name from store_design_template where store_id="&Store_Id&" and template_id="&Template_Id
		set rs_select=conn_store.execute(sql_select)
		templ_name=rs_select("template_name")
		set rs_select=nothing


	'	sTitle = "Manage "&design_area_name&" Area"

		sTitle = "Edit Template "&design_area_name&" Properties"
		sFullTitle = sFullTitleStart & "Edit "&design_area_name&" Properties"
		sCommonName = design_area_name&" Properties"
                sFormAction = "layout_settings.asp"
		thisRedirect = "manage_object.asp"
		sFormName ="objForm"
		sMenu = "design"
		addPicker=1
		Calendar=1
		createHead thisRedirect


		sql_select="select "&Design_Area_BG&" as Design_Area_BG_Val ,"&Design_Area_Border&" as Design_Area_Border_Val ,"&Design_Area_Height&" as Design_Area_Height_Val ,"&Design_Area_Width&" as Design_Area_Width_Val,"&Design_Area_Border_Color&" as Design_Area_Border_Color_Val   from store_design_template where Store_Id="&Store_Id&" and template_id="&Template_Id
	    set rs_select=conn_store.execute(sql_select)
		if not rs_select.EOF then
			Design_Area_BG_Val=rs_select("Design_Area_BG_Val")
			Design_Area_Border_Val=rs_select("Design_Area_Border_Val")
			Design_Area_Height_Val=rs_select("Design_Area_Height_Val")
			Design_Area_Width_Val=rs_select("Design_Area_Width_Val")
			Design_Area_Border_Color_Val=rs_select("Design_Area_Border_Color_Val")

		else
			fn_error "Template does not exist"
		end if

		
		%>
				<TR bgcolor='#FFFFFF'>
					 <td width="100%" colspan="3" height="11">
						 <input type="button" OnClick="javaScript:self.location='layout_design.asp?Id=<%=Template_Id%>'" class="Buttons" value="Template Detail" name="Design_Layout_Page"></td>
				</tr>
		  
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Background Color or file</td>
			<td width="70%" class="inputvalue">
				<input size=60 type="text" name="Design_Area_BG" value="<%=Design_Area_BG_Val%>">
				<a href="javascript:goColorPicker('Design_Area_BG')"><img src="images/bgcol.gif" width="22" height="22" border="0"></a> <a href="javascript:goImagePicker('Design_Area_BG')"><img src="images/image.gif" width="22" height="22" border="0"></a> <a class="link" href="JavaScript:goFileUploader('Design_Area_BG');"><img src="images/folderup.gif" width="23" height="22" border="0"></a>
			 <% small_help "Area Design BG" %>
			 </td>
		  </tr>

			
				  
		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Height</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=4 name="Design_Area_Height" value="<% = Design_Area_Height_Val %>">
				<% small_help "Area Height" %>
				</td>
		  </tr>

		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Width</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=4 name="Design_Area_Width" value="<% = Design_Area_Width_Val %>">
				<% small_help "Area Width" %>
				</td>
		  </tr>
			
			
		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=1 maxlength=3 name="Design_Area_Border" value="<% = Design_Area_Border_Val %>" onKeyPress="return goodchars(event,'0123456789.')">
				<% small_help "Area Border" %>
				</td>
		  </tr>

			
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Design_Area_Border_Color" id="Design_Area_Border_Color" value="<%= Design_Area_Border_Color_Val %>" onKeyUp="ChangeColor1('Design_Area_Border_Color','Design_Area_Border_Color_ValPreview');" onFocus="ChangeColor1('Design_Area_Border_Color','Design_Area_Border_Color_ValPreview'); " onBlur="ChangeColor1('Design_Area_Border_Color','Design_Area_Border_Color_ValPreview');">

				<% 'COLOR PREVIEW BOX
				if len(Design_Area_Border_Color_Val) <> 0 then
				%>
					<input type="text"  size=1 id="Design_Area_Border_Color_ValPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= Design_Area_Border_Color_Val %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Design_Area_Border_Color')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
				<% small_help "Area Border Color" %>	
			</td>
		  </tr>

			
			<input type="hidden" value="<%= ar %>" name="Area">
			<input type="hidden" value="<%= Template_Id %>" name="Template_Id">

		   <%createFoot thisRedirect, 1%>
			

<%
	elseif Request.QueryString("op") <> "" then


		op=Request.QueryString("op")
		Template_Id=Request.QueryString("tid")
		
		sql_select="select template_name from store_design_template where store_id="&Store_Id&" and template_id="&Template_Id
		set rs_select=conn_store.execute(sql_select)
		templ_name=rs_select("template_name")
		set rs_select=nothing


		if op="add" then
			Design_Area=Request.QueryString("area")

			select case Design_Area
				case 1
					Design_Area_Name="Top"
				case 2
					Design_Area_Name="Left"
				case 3
					Design_Area_Name="Right"
				case 4
					Design_Area_Name="Bottom"
				case 5
					Design_Area_Name="Center Top"
				case 6
					Design_Area_Name="Center Bottom"
			end select

		sTitle = "Add Template "&Design_Area_Name& " Object"
                sFullTitle = sFullTitleStart & "Add "&Design_Area_Name&" Object"
		sName = "Manage_Object"
		sFormAction = "layout_settings.asp"
		thisRedirect = "manage_object.asp"
		sFormName ="Manage_Object"
		sMenu = "design"
		addPicker=1
		sSubmitName = "Update_Obj"
		createHead thisRedirect
		


	%>
	
	
	<TR bgcolor='#FFFFFF'> 
		<td width="30%" class="inputname">Object</td>
		<td width="70%" class="inputvalue"><select name="Object_Type" size="1">
			<option value="Object_Type">Select Object</option>

			<option value="Banner">Banner</option>
			<option value="Date">Date</option>
			<option value="Department Select Box">Department Select Box</option>			
			<option value="HR">Horizontal Rule</option>
			<option value="HTML Text">HTML Text</option>
			<option value="Image">Image</option>
			<option value="Link">Link</option>
			<option value="Login Box">Login Box</option>			
			<option value="Nav Buttons">Nav Buttons</option>
			<option value="Nav Links">Nav Links</option>
			<option value="Newsletter Signup">Newsletter Signup Box</option>
			<option value="Search Box">Search Box</option>
			<option value="Small Cart">Small Cart</option>
			<option value="Space">Space</option>

			
		  </select></td>
	  </tr>

		<input type="hidden" name="Design_Area" value="<%= Design_Area%>">
		<input type="hidden" name="op" value="add">
		<input type="hidden" name="Template_Id" value="<%=Template_Id%>">

	   <%createFoot thisRedirect, 1%>

<%


		elseif  op="edit" then

	

		 if request.querystring("Id")<> "" then
			Object_Id=Request.QueryString("Id")
            if Object_Id="New" then
			    sql_select="select top 1 * from store_design_objects where Object_Type='"&request.QueryString("Object_Type")&"' and Store_Id="&Store_Id&" and template_id="&Template_Id&" order by object_id desc"
			else
			    sql_select="select * from store_design_objects where Object_Id="&Object_Id&" and Store_Id="&Store_Id&" and template_id="&Template_Id
			end if
			set rs_select=conn_store.execute(sql_select)
			
			if rs_select. EOF then
				fn_error "Design Object does not exist"
			end if
            Object_Id=rs_select("Object_Id")
			Object_Type=rs_select("Object_Type")
			Design_Area=rs_select("Design_Area")	
			
			select case Design_Area
				case 1
					Design_Area_Name="Top"
				case 2
					Design_Area_Name="Left"
				case 3
					Design_Area_Name="Right"
				case 4
					Design_Area_Name="Bottom"
				case 5
					Design_Area_Name="Center Top"
				case 6
					Design_Area_Name="Center Bottom"
			end select

			Object_Area_Height=rs_select("Object_Area_Height")
			Object_Area_Width=rs_select("Object_Area_Width") 
			Object_Alignment=rs_select("Object_Alignment")
			if Object_Alignment = ""  or Object_Alignment = null then
				Object_Alignment="Center"
			end if
		
		sTitle = "Edit Template "&Design_Area_Name&" "&Object_Type
                sFullTitle = sFullTitleStart & "Edit "&Design_Area_Name&" "&Object_Type
		sName = "Manage_Object"
		sFormAction = "layout_settings.asp"
		thisRedirect = "manage_object.asp"
		sFormName ="Manage_Object"
		sMenu = "design"
		addPicker=1
		sSubmitName = "Update_Obj"
		createHead thisRedirect
		
		%>

		
		 
			<TR bgcolor='#FFFFFF'>
					 <td width="100%" colspan="3" height="11">
						 <input type="button" OnClick="javaScript:self.location='layout_design.asp?Id=<%=Template_Id%>'" class="Buttons" value="Template Detail" name="Design_Layout_Page"></td>
			</tr>
	
	<% select case Object_Type %>

		<% case "Image","Banner" %>

	   <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Image Preview</td>
			<td width="70%" class="inputvalue" colspan=2><img name="Image" src="<% =Site_Name&"images/"%><% =rs_select("Image_Path")%>" alt="<% = rs_select("Image_Alt")%>"	border="0">
		</tr>

	   <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Image Path</td>
			<td width="70%">
				<input type="text" size=60 maxlength=100 name="Image_Path" value="<%=rs_select("Image_Path")%>">
				<input type="hidden" name="Image_Path_C" value="Op|String|0|100|||Image Path">
				<a href="javascript:goImagePicker('Image_Path')"><img src="images/image.gif" width="22" height="22" border="0"></a> <a class="link" href="JavaScript:goFileUploader('Image_Path');"><img src="images/folderup.gif" width="23" height="22" border="0"></a>
			<% small_help "Image Path"%>
			</td>
		  </tr>

  		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Alternate Text</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=60 maxlength=50 name="Image_Alt" value="<% = rs_select("Image_Alt")%>">
				<input type="hidden" name="Image_Alt_C" value="Op|String|0|50|||Alternate Text">
						<% small_help "Image Alternate Text" %>
				</td>
		  </tr>

		

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border</td>
			<td width="70%" class="inputvalue"><select name="Border_Size" size="1">
				<option value="<% = rs_select("Border_Size")%>" selected><% = rs_select("Border_Size")%></option>
				<option value="0">0</option>
				<option value="1">1</option>
				<option value="2">2</option>
				<option value="3">3</option>
				<option value="4">4</option>
				<option value="5">5</option>
			  </select>
			  <% small_help "Image Border" %>
			  </td>
		  </tr>


		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Vertical Spacing</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=4 name="Image_VSpace" value="<% = rs_select("Image_VSpace")%>">
				<% small_help "Image VSpace" %>
				</td>
		  </tr>

		
		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Horizontal Spacing</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=4 name="Image_HSpace" value="<% = rs_select("Image_HSpace")%>">
				<% small_help "Image HSpace" %>
				</td>
		  </tr>
			
	
		<%	if Object_Type = "Banner" then %>
		
		  <TR bgcolor='#FFFFFF'> 
				<td width="30%" class="inputname">Protocol</td>
				<td width="70%" class="inputvalue">
					<select name="Link_Protocol" size="1">
					<%if rs_select("Link_Protocol") <> "" then%>
						<option value="<%=rs_select("Link_Protocol")%>" selected><%= rs_select("Link_Protocol")%></option>
					<%end if%>
					  <option value="http://">http://</option>
					  <option value="https://">https://</option>
					  <option value="mailto:">mailto:</option>
  					  <option value="ftp://">ftp://</option>
					</select>
					<% small_help "Banner Protocol"%>			
				
				</td>
		  </tr>
					
		  <TR bgcolor='#FFFFFF'> 
				<td width="30%" class="inputname">Target</td>
				<td width="70%" class="inputvalue">
					<select name="Link_Target" size="1">
					<%if rs_select("Link_Target") <> "" then%>
						  <option value="<%=rs_select("Link_Target")%>" selected><%= rs_select("Link_Target")%></option>
					 <%end if%>
					  <option value="self">self</option>
					  <option value="blank">blank</option>
					  <option value="parent">parent</option>
					</select>
					<% small_help "Banner Target"%>			
				
				</td>
		  </tr>
			  

		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">URL</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=50 name="Link_URL" value="<% = rs_select("Link_URL")%>">
				<% small_help "Banner URL" %>
				</td>
		  </tr>


		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Text</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=50 name="Link_Text" value="<% = rs_select("Link_Text")%>">
				<% small_help "Banner Text" %>
				</td>
		  </tr>


		<%end if%>	

			

		<% case "Space" %>

		

		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Background Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="BG" id="BG" value="<%=rs_select("BG")%>" onKeyUp="ChangeColor1('BG','BGPreview');" onFocus="ChangeColor1('BG','BGPreview'); " onBlur="ChangeColor1('BG','BGPreview');">


				<% 'COLOR PREVIEW BOX
				if len(rs_select("BG")) <> 0 then
				%>
					<input type="text"  size=1 id="BGPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("BG") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('BG')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
					<% small_help "Space BG Color" %>
				</td> 
		</tr>


		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Border_Color" id="Border_Color" value="<%=rs_select("Border_Color")%>" onKeyUp="ChangeColor1('Border_Color','Border_ColorPreview');" onFocus="ChangeColor1('Border_Color','Border_ColorPreview'); " onBlur="ChangeColor1('Border_Color','Border_ColorPreview');">


				<% 'COLOR PREVIEW BOX
				if len(rs_select("Border_Color")) <> 0 then
				%>
					<input type="text"  size=1 id="Border_ColorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Border_Color") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Border_Color')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
					<% small_help "Space Border Color" %>
				</td>
		</tr>

		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Size</td>
			<td width="70%" class="inputvalue">
			<input type="text" name="Border_Size" value="<%=rs_select("Border_Size")%>" size="1" maxlength="3" onKeyPress="return goodchars(event,'0123456789.')">
					<% small_help "Space Border Size" %>
				</td>
		  </tr>



		<% case "Link" %>
				
 		  <TR bgcolor='#FFFFFF'> 
				<td width="30%" class="inputname">Protocol</td>
				<td width="70%" class="inputvalue">
					<select name="Link_Protocol" size="1">
					<%if rs_select("Link_Protocol") <> "" then%>
						  <option value="<%=rs_select("Link_Protocol")%>" selected><%= rs_select("Link_Protocol")%></option>
					<%end if%>
					  <option value="http://">http://</option>
					  <option value="https://">https://</option>
					  <option value="mailto:">mailto:</option>
  					  <option value="ftp://">ftp://</option>
					</select>
					<% small_help "Link Protocol"%>			
				
				</td>
		  </tr>

		  
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Target</td>
			<td width="70%" class="inputvalue">
				<select name="Link_Target" size="1">
			  	<%if rs_select("Link_Target") <> "" then%>
					  <option value="<%=rs_select("Link_Target")%>" selected><%= rs_select("Link_Target")%></option>
				<%end if%>
				  <option value="self">self</option>
				  <option value="blank">blank</option>
				  <option value="parent">parent</option>
				</select>
				<% small_help "Link Target"%>			
			</td>
		  </tr>

		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">URL</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=50 name="Link_URL" value="<% = rs_select("Link_URL")%>">
				<% small_help "URL" %>
				</td>
		  </tr>


		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Text</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=50 name="Link_Text" value="<% = rs_select("Link_Text")%>">
				<% small_help "Link Text" %>
				</td>
		  </tr>

		
		
		<% case "Nav Links" %>
				
	  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Position</td>
			<td width="70%" class="inputvalue">
				<select name="Nav_Position" size="1">
				<% if rs_select("Nav_Position") <> "" then%>
					  <option value="<%=rs_select("Nav_Position")%>"><%=rs_select("Nav_Position")%></option>
				 <%end if%>
					<option value="Horizontal">Horizontal</option>				 
					<option value="Vertical">Vertical</option>
				</select>
				<% small_help "Nav Links Position" %>	
			
			  </td>
		  </tr>

		

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Spacing</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=3 name="Nav_Spacing" value="<% = rs_select("Nav_Spacing")%>">
				pixels 
				<% small_help "Nav Links Spacing" %>
				</td>
		  </tr>
    		<TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Menu</td>
			<td width="70%" class="inputvalue"> 
				<select name="Nav_Menu_Number" size="1">
				<option value="<%=rs_select("Nav_Menu_Number")%>" selected><%=rs_select("Nav_Menu_Number")%></option>
				  <option value=1>1</option>
				  <option value=2>2</option>
				  <option value=3>3</option>
				  <option value=4>4</option>
				  <option value=5>5</option>
				</select>
				<% small_help "Nav Menu" %>
			  </td>
		  </tr>
		  <TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Font</td>
			<td width="70%" class="inputvalue"> 
				<select name="Nav_Font_Face" size="1">
				<option value="<%=rs_select("Nav_Font_Face")%>" selected><%= rs_select("Nav_Font_Face")%></option>
				  <option value="Arial">Arial</option>
				  <option value="Verdana">Verdana</option>
				  <option value="Tahoma">Tahoma</option>
				  <option value="Times New Roman">Times New Roman</option>
				</select>
				<% small_help "Nav Links Font" %>
			  </td>
		  </tr>
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Font Size</td>
			<td width="70%" class="inputvalue">
				
				<input type="text" name="Nav_Font_Size" value="<%=rs_select("Nav_Font_Size")%>" size=2 maxlength=3 onKeyPress="return goodchars(event,'0123456789.')">
				<% small_help "Nav Links Font Size" %>
			  </td>
		  </tr>
		
		  <TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Bold</td>
			<td width="70%" class="inputvalue">
						<input class="image" name="Nav_Bold" type="checkbox" value=-1 
					<%	if rs_select("Nav_Bold") =-1 then %>
							 checked 
					 <%  end if%>
					 >
					<% small_help "Nav Links Bold" %>
				</td>
			</tr>

		  <TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Italic</td>
			<td width="70%" class="inputvalue">
						<input class="image" name="Nav_Italic" type="checkbox" value=-1 
					<%	if rs_select("Nav_Italic") =-1 then %>
							 checked 
					 <%  end if%>
					 >
					<% small_help "Nav Links Italic" %>
				</td>
			</tr>

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Link Text Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Nav_Font_Color" id = "Nav_Font_Color"value="<%=rs_select("Nav_Font_Color")%>" onKeyUp="ChangeColor1('Nav_Font_Color','Nav_Font_ColorPreview');" onFocus="ChangeColor1('Nav_Font_Color','Nav_Font_ColorPreview'); " onBlur="ChangeColor1('Nav_Font_Color','Nav_Font_ColorPreview');">


				<% 'COLOR PREVIEW BOX
				if len(rs_select("Nav_Font_Color")) <> 0 then
				%>
					<input type="text"  size=1 id="Nav_Font_ColorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Nav_Font_Color") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Nav_Font_Color')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
				<% small_help "Nav Links Text Color" %>	
			</td>
		  </tr>
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Visited Link Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Color_1" id ="Color_1" value="<%=rs_select("Color_1")%>"  onKeyUp="ChangeColor1('Color_1','Color_1Preview');" onFocus="ChangeColor1('Color_1','Color_1Preview'); " onBlur="ChangeColor1('Color_1','Color_1Preview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("Color_1")) <> 0 then
				%>
					<input type="text"  size=1 id="Color_1Preview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Color_1") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Color_1')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
					<% small_help "Nav Links Visited Color" %>	
				</td>
		  </tr>
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Hover Link Color</font></td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="color_2" id="color_2" value="<%=rs_select("color_2")%>" onKeyUp="ChangeColor1('color_2','color_2Preview');" onFocus="ChangeColor1('color_2','color_2Preview'); " onBlur="ChangeColor1('color_2','color_2Preview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("color_2") ) <> 0 then
				%>
					<input type="text"  size=1 id="color_2Preview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("color_2") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('color_2')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
				<% small_help "Nav Links Hover Color" %>
				</td>
		  </tr>


		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Background Color or file</td>
			<td width="70%" class="inputvalue">
				<input type="text" name="BG" value="<%=rs_select("BG")%>">
				<a href="javascript:goColorPicker('BG')"><img src="images/bgcol.gif" width="22" height="22" border="0"> </a>
				 <a href="javascript:goImagePicker('BG')"><img src="images/image.gif" width="22" height="22" border="0"></a> <a class="link" href="JavaScript:goFileUploader('BG');"><img src="images/folderup.gif" width="23" height="22" border="0"></a>
			</td>
				<% small_help "Nav Links BG Color" %>
		  </tr>

		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Size</td>
			<td width="70%" class="inputvalue">
				<input type="text"  name="Border_Size" value="<% = rs_select("Border_Size")%>" size=1 maxlength="3" onKeyPress="return goodchars(event,'0123456789.')">
					pixels
				<% small_help "Nav Links Border Size" %>
				</td>
		  </tr>
		
		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Border_Color" id="Border_Color" value="<%=rs_select("Border_Color")%>" onKeyUp="ChangeColor1('Border_Color','Border_ColorPreview');" onFocus="ChangeColor1('Border_Color','Border_ColorPreview'); " onBlur="ChangeColor1('Border_Color','Border_ColorPreview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("Border_Color")) <> 0 then
				%>
					<input type="text"  size=1 id="Border_ColorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Border_Color") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>



				<a href="javascript:goColorPicker('Border_Color')"><img src="images/bgcol.gif" width="22" height="22" border="0"> </a>
			</td>
				<% small_help "Nav Links Border BG Color" %>
		  </tr>

		 <TR bgcolor='#FFFFFF'>	
			<td width="30%" class="inputname">Links Seperator</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=2 name="Nav_Links_Seperator" value="<% = rs_select("Nav_Links_Seperator")%>">
				<% small_help "Nav Links Seperator" %>
				</td>
		  </tr>

		
		<% case "Nav Buttons" %>

		

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Position</td>
			<td width="70%" class="inputvalue">
				<select name="Nav_Position" size="1">
				<% if rs_select("Nav_Position") <> "" then%>
					  <option value="<%=rs_select("Nav_Position")%>"><%=rs_select("Nav_Position")%></option>
				 <%end if%>
				  <option value="Vertical">Vertical</option>
				  <option value="Horizontal">Horizontal</option>
				</select>
				<% small_help "Nav Buttons Position" %>	
			
			  </td>
		  </tr>


		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Background Color or file</td>
			<td width="70%" class="inputvalue">
				<input type="text" name="BG" value="<%=rs_select("BG")%>">
				<a href="javascript:goColorPicker('BG')"><img src="images/bgcol.gif" width="22" height="22" border="0"> </a>
				 <a href="javascript:goImagePicker('BG')"><img src="images/image.gif" width="22" height="22" border="0"></a><a class="link" href="JavaScript:goFileUploader('BG');"><img src="images/folderup.gif" width="23" height="22" border="0"></a>
			</td>
				<% small_help "Nav Buttons BG" %>
		  </tr>

		

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Spacing</td>
			<td width="70%" class="inputvalue"><input type="text" name="Nav_Spacing" size="3" value="<%=rs_select("Nav_Spacing")%>">
				pixels 
					<% small_help "Nav Buttons Spacing" %>	
				</td>
		  </tr>
		  <TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Menu</td>
			<td width="70%" class="inputvalue"> 
				<select name="Nav_Menu_Number" size="1">
				<option value="<%=rs_select("Nav_Menu_Number")%>" selected><%= rs_select("Nav_Menu_Number")%></option>
				  <option value=1>1</option>
				  <option value=2>2</option>
				  <option value=3>3</option>
				  <option value=4>4</option>
				  <option value=5>5</option>
				</select>
				<% small_help "Nav Menu" %>
			  </td>
		  </tr>

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Button Width</td>
			<td width="70%" class="inputvalue"><input type="text" name="Object_Width" size="4" value="<%=rs_select("Object_Width")%>">
			<% small_help "Nav Buttons Width" %>	
				</td>
		  </tr>
		  
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Button Height</td>
			<td width="70%" class="inputvalue"><input type="text" name="Object_Height" size="4" value="<%=rs_select("Object_Height")%>">
			<% small_help "Nav Buttons Height" %>	
				</td>
		  </tr>
		  

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Text Align</td>
			<td width="70%" class="inputvalue">
				<select name="Alignment" size="1">
				  <option value="<%=rs_select("Alignment")%>"><%=rs_select("Alignment")%></option>
				  <option value="Center">Center</option>
				  <option value="Right">Right</option>
				  <option value="Left">Left</option>
				</select>
				<% small_help "Nav Buttons Text Align" %>	
			
			  </td>
		  </tr>

		  
		  
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Font</td>
			<td width="70%" class="inputvalue">
				<select name="Nav_Font_Face" size="1">
				<option value="<%=rs_select("Nav_Font_Face")%>"><%=rs_select("Nav_Font_Face")%></option>
				  <option value="Arial">Arial</option>
				  <option value="Verdana">Verdana</option>
				  <option value="Tahoma">Tahoma</option>
				  <option value="Times New Roman">Times New Roman</option>
				</select>
				<% small_help "Nav Buttons Font" %>	
			  </td>
		  </tr>
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Font Size</td>
			<td width="70%" class="inputvalue">
			
				<input type="text" name="Nav_Font_Size" value="<%=rs_select("Nav_Font_Size")%>" size=2 maxlength=3 onKeyPress="return goodchars(event,'0123456789.')">
				<% small_help "Nav Buttons Font Size" %>
			  </td>
		  </tr>

		  <TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Bold</td>
			<td width="70%" class="inputvalue">
						<input class="image" name="Nav_Bold" type="checkbox" value=-1 
					<%	if rs_select("Nav_Bold") =-1 then %>
							 checked 
					 <%  end if%>
					 >
					<% small_help "Nav Buttons Bold" %>
				</td>
			</tr>

		  <TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Italic</td>
			<td width="70%" class="inputvalue">
						<input class="image" name="Nav_Italic" type="checkbox" value=-1 
					<%	if rs_select("Nav_Italic") =-1 then %>
							 checked 
					 <%  end if%>
					 >
					<% small_help "Nav Buttons Italic" %>
				</td>
			</tr>


		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Button Text Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="nav_font_color" id="nav_font_color" value="<%=rs_select("nav_font_color")%>" onKeyUp="ChangeColor1('nav_font_color','nav_font_colorPreview');" onFocus="ChangeColor1('nav_font_color','nav_font_colorPreview'); " onBlur="ChangeColor1('nav_font_color','nav_font_colorPreview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("nav_font_color")) <> 0 then
				%>
					<input type="text"  size=1 id="nav_font_colorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("nav_font_color") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('nav_font_color')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
			</td>
			<% small_help "Nav Buttons Text Color" %>
		  </tr>

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Border_Color" id="Border_Color" value="<%=rs_select("Border_Color")%>"				onKeyUp="ChangeColor1('Border_Color','Border_ColorPreview');" onFocus="ChangeColor1('Border_Color','Border_ColorPreview'); " onBlur="ChangeColor1('Border_Color','Border_ColorPreview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("Border_Color")) <> 0 then
				%>
					<input type="text"  size=1 id="Border_ColorPreview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Border_Color") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Border_Color')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
					<% small_help "Nav Buttons Border Color" %>
				</td>
		  </tr>

		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Border Size</td>
			<td width="70%" class="inputvalue">
			<select name="Border_Size" size="1">
				<option value="<%=rs_select("Border_Size")%>"><%=rs_select("Border_Size")%></option>
				  <option value="0">0</option>
				  <option value="1">1</option>
				  <option value="2">2</option>
				  <option value="3">3</option>
				  <option value="4">4</option>
				</select>
					<% small_help "Nav Buttons Border Size" %>
				</td>
		  </tr>




		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Button Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Color_1" id="Color_1" value="<%=rs_select("Color_1")%>" onKeyUp="ChangeColor1('Color_1','Color_1Preview');" onFocus="ChangeColor1('Color_1','Color_1Preview'); " onBlur="ChangeColor1('Color_1','Color_1Preview');">


				<% 'COLOR PREVIEW BOX
				if len(rs_select("Color_1")) <> 0 then
				%>
					<input type="text"  size=1 id="Color_1Preview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Color_1") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Color_1')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
					<% small_help "Nav Buttons Color" %>
				</td>
		  </tr>
		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Active Button Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="color_2" id ="color_2" value="<%=rs_select("color_2")%>" onKeyUp="ChangeColor1('color_2','color_2Preview');" onFocus="ChangeColor1('color_2','color_2Preview'); " onBlur="ChangeColor1('color_2','color_2Preview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("color_2") ) <> 0 then
				%>
					<input type="text"  size=1 id="color_2Preview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("color_2") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('color_2')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
				
				<% small_help "Nav Buttons Active Button Color" %>
				</td>
		  </tr>


		<% case "HTML Text" %>
			<TR bgcolor='#FFFFFF'> 
				<td class="inputname" colspan=2>Text

				       
						<%= create_editor ("HTML_Text",rs_select("HTML_Text"),"") %>

                                                <% small_help "HTML Text" %>
				 </td>
			  </tr>


		<% case "HR" %>
		
		

		  <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Color</td>
			<td width="70%" class="inputvalue">
				<input size=7 maxlength=7 type="text" name="Color_1" id="Color_1" value="<%=rs_select("Color_1")%>" onKeyUp="ChangeColor1('Color_1','Color_1Preview');" onFocus="ChangeColor1('Color_1','Color_1Preview'); " onBlur="ChangeColor1('Color_1','Color_1Preview');">

				<% 'COLOR PREVIEW BOX
				if len(rs_select("Color_1") ) <> 0 then
				%>
					<input type="text"  size=1 id="Color_1Preview" disabled style="border:1;border-style:solid;border-color:#000000;background-Color:#<%= rs_select("Color_1") %>" >		
						
				<%
				end if
				 'COLOR PREVIEW BOX END
				 %>

				<a href="javascript:goColorPicker('Color_1')"><img src="images/fontcolor.gif" width="22" height="22" border="0"></a>
					<% small_help "HR Color" %>
				</td>
		  </tr>

		
		<% case else %>
		
		<% end select %>
		
		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Height</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=4 name="Object_Area_Height" value="<% = Object_Area_Height %>">
				<% small_help "Object Height" %>
				</td>
		  </tr>

		 <TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Width</td>
			<td width="70%" class="inputvalue">
				<input type="text" size=4 name="Object_Area_Width" value="<% =  Object_Area_Width %>">
				<% small_help "Object Width" %>
				</td>
		  </tr>
			



		<TR bgcolor='#FFFFFF'> 
			<td width="30%" class="inputname">Alignment</td>
			<td width="70%" class="inputvalue">
				<select name="Object_Alignment" size="1">
					<option value="<%= Object_Alignment%>" selected><%= Object_Alignment %></option>
		  		    <option value="Left">Left</option>
				    <option value="Right">Right</option>
				    <option value="Center">Center</option>
				</select>
				<% small_help "Object Alignment" %>
			  </td>
		  </tr>

		<TR bgcolor='#FFFFFF'>
			<td width="30%" class="inputname">Design Area</td>
			<td width="70%" class="inputvalue"><select name="Design_Area" size="1">
				<option value="<%=rs_select("Design_Area")%>"><%= Design_Area_Name%></option>
				<option value="1">Top</option>
				<option value="2">Left</option>
				<option value="3">Right</option>
				<option value="4">Bottom</option>
				<option value="5">Center Top</option>
				<option value="6">Center Bottom</option>
			  </select>
			   <% small_help "Design Area" %>
			  </td>
		  </tr>
		  <TR bgcolor='#FFFFFF'>
		  	<td width="30%" class="inputname">Next Object</td>
		  	<% if rs_select("same_line") = 0 then 
		  	      radio_0="checked"
		  	      radio_1=""
                           else
                              radio_1="checked"
                              radio_0=""
                           end if %>
			<td width="70%" class="inputvalue"><input type=radio <%= radio_1 %> value=1 name=Same_Line>On same line<BR>
                        <input type=radio <%= radio_0 %> value=0 name=Same_Line>On new line
			   <% small_help "Same Line" %>
			  </td>
		  </tr>

			<input type="hidden" value="<%= Object_Id %>" name="Object_Id">
			<input type="hidden" value="<%= Object_Type %>" name="Object_Type">
			<input type="hidden" name="op" value="edit">
			<input type="hidden" name="Template_Id" value="<%=Template_Id%>">

		<% createFoot thisRedirect, 1 %>

		

	<%

		end if
	
	end if	'add/edit check

end if	' area edit check

end if	'del check



%>
