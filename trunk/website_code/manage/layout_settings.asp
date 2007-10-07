<!--#include file="Global_Settings.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include file="pagedesign.asp"-->
<%	

if (Request.Form <> "" or Request.QueryString <> ""  )then
    if Request("Gen") =1 then
        ' CAME FROM APPLY TEMPLATE AND PREVIEW TEMPLATE
	    Template_Id=Request("Template_Id")
        Preview_Id=Request("Preview_Id")
        call sub_template_select(Template_Id,Preview_Id)
        response.redirect "layout_design.asp?Id="&Template_Id
	end if

    if Request.Form("Design_Area_BG") <> "" or Request.Form("Area") <> "" then
	    'CAME FROM MANAGE_OBJECT FOR DESIGN AREA PROPERTIES
		Template_Id=Request.Form("Template_Id")
		Design_Area_BG=Request.Form("Design_Area_BG")
		Design_Area_Border=Request.Form("Design_Area_Border")
		Design_Area_Height=Request.Form("Design_Area_Height")
		Design_Area_Width=Request.Form("Design_Area_Width")
		Design_Area_Border_Color=Request.Form("Design_Area_Border_Color")
		Template_Id=Request.Form("Template_Id")

		if Design_Area_Border_Color="" then
			Design_Area_Border_Color="FFFFFF"
		end if
		Area=Request.Form("Area")

		select case Area
			case 1:
			    sAreaName="Top"
			case 2:
			    sAreaName="Left"
			case 3:
			    sAreaName="Right"
			case 4:
			    sAreaName="Bottom"
			case 5:
			    sAreaName="Cen_Top"
			case 6:
			    sAreaName="Cen_Bot"
			case 7:
			    sAreaName="Cen_Cen"
		end select
		where_update="Design_"&sAreaName&"_BG='"&Design_Area_BG&"' ,Design_"&sAreaName&"_Border='"&Design_Area_Border&"', Design_"&sAreaName&"_Height='"&Design_Area_Height&"', Design_"&sAreaName&"_Width='"&Design_Area_Width&"', Design_"&sAreaName&"_Border_Color='"&Design_Area_Border_Color&"' " 
			
		sql_update="update store_design_template set "&where_update&" where store_id="&Store_Id&" and template_id="&Template_Id
		conn_store.execute(sql_update)

		response.redirect "layout_design.asp?Id="&Template_Id
	end if
	
	if Request.Form("Left_Order") <> "" then
	    ' CAME FROM REORDERING OF DESIGN OBJECTS ON THE MAIN LAYOUT PAGE
		Template_Id=Request.Form("Template_Id")
		Left_Order=Request.Form("Left_Order")
		Left_Order_Arr=split(Left_Order, ",")

		for i=0 to Ubound(Left_Order_Arr)
		    sql_update="update store_design_objects set view_order="&	i+1&" where Object_Id="&Left_Order_Arr(i)&" and store_id="&Store_Id&" and template_id="&Template_Id
		    conn_store.execute(sql_update)
		next

		response.redirect "layout_design.asp?Id="&Template_Id
	end if

	if Request.Form("op") <> "" then
        'CAME FROM MANAGE_OBJECT FOR DESIGN OBJECTS 
		op=Request.Form("op")
		Template_Id=Request.Form("Template_Id")
		Design_Area=Request.Form("Design_Area")
		Same_line=Request.Form("Same_line")
        Object_Type=trim(Request.Form("Object_Type"))
		
		if op = "edit" then
			Object_Id=Request.Form("Object_Id")
			View_Order=Request.Form("View_Order")
		    Object_Alignment=trim(Request.Form("Object_Alignment"))
		    Object_Area_Height=trim(Request.Form("Object_Area_Height"))
		    Object_Area_Width=trim(Request.Form("Object_Area_Width"))
    		
		    select case Object_Type
			    case "Image"
		            Image_Path=Request.Form("Image_Path")
		            Image_Alt=Request.Form("Image_Alt")
		            Object_Width=Request.Form("Object_Width")
		            Object_Height=Request.Form("Object_Height")
		            Alignment=Request.Form("Alignment")
		            Border_Size=Request.Form("Border_Size")
		            Image_VSpace=Request.Form("Image_VSpace")
		            Image_HSpace=Request.Form("Image_HSpace")
		            call sub_set_defaults()
					sql_where="Object_Height='"&Object_Height&"' ,Object_Width='"&Object_Width&"' ,Image_Path='"&Image_Path&"' , Alignment='"&Alignment&"', Image_Alt='"&Image_Alt&"' ,Border_Size='"&Border_Size&"' ,Image_VSpace='"&Image_VSpace&"' , Image_HSpace='"&Image_HSpace&"',"
			    case "Banner"
				    Image_Path=Request.Form("Image_Path")
				    Image_Alt=Request.Form("Image_Alt")
				    Object_Width=Request.Form("Object_Width")
				    Object_Height=Request.Form("Object_Height")
				    Alignment=Request.Form("Alignment")
				    Border_Size=Request.Form("Border_Size")
				    Image_VSpace=Request.Form("Image_VSpace")
				    Image_HSpace=Request.Form("Image_HSpace")
				    Link_URL=Request.Form("Link_URL")
				    Link_Target=Request.Form("Link_Target")
				    Link_Text=Request.Form("Link_Text")
				    Link_Protocol=Request.Form("Link_Protocol")
				    call sub_set_defaults()
				    sql_where="Object_Height='"&Object_Height&"' ,Object_Width='"&Object_Width&"' ,Image_Path='"&Image_Path&"' , Alignment='"&Alignment&"', Image_Alt='"&Image_Alt&"' ,Border_Size='"&Border_Size&"' ,Image_VSpace='"&Image_VSpace&"' , Image_HSpace='"&Image_HSpace&"',Link_URL='"&Link_URL&"',Link_Target='"&Link_Target&"',Link_Text='"&Link_Text&"',Link_Protocol='"&Link_Protocol&"',"
			    case "Link"
				    Link_URL=Request.Form("Link_URL")
				    Link_Target=Request.Form("Link_Target")
				    Link_Text=Request.Form("Link_Text")
				    Link_Protocol=Request.Form("Link_Protocol")
				    call sub_set_defaults()
				    sql_where="Link_URL='"&Link_URL&"',Link_Target='"&Link_Target&"',Link_Text='"&Link_Text&"',Link_Protocol='"&Link_Protocol&"',"
			    case "Space"
				    Object_Height=Request.Form("Object_Height")
				    Object_Width=Request.Form("Object_Width")
				    Border_Size=Request.Form("Border_Size")
				    Border_Color=Request.Form("Border_Color")
				    BG=Request.Form("BG")	
				    if Border_Size = "" or Border_Size = null then
					    Border_Size="0"
				    end if
				    if Border_Color = "" or Border_Color = null then
					    Border_Color="FFFFFF"
				    end if
				    if Object_Alignment = "" or Object_Alignment = null then
					    Object_Alignment="Center"
				    end if
				    call sub_set_defaults()
				    sql_where="Object_Height='"&Object_Height&"',Object_Width='"&Object_Width&"',Border_Size='"&Border_Size&"',Border_Color='"&Border_Color&"',BG='"&BG&"',"
			    case "Nav Links"
				    Nav_Menu_Number=Request.Form("Nav_Menu_Number")
				    Nav_Links_Alignment=Request.Form("Nav_Links_Alignment")
				    Nav_Spacing=Request.Form("Nav_Spacing")
				    Nav_Font_Face=Request.Form("Nav_Font_Face")
				    Nav_Font_Size=Request.Form("Nav_Font_Size")
				    Nav_Font_Color=Request.Form("Nav_Font_Color")
				    color_1=Request.Form("Color_1")
				    color_2=Request.Form("color_2")
				    Nav_Position=Request.Form("Nav_Position")
				    BG=Request.Form("BG")
				    Border_Color=Request.Form("Border_Color")
				    Border_Size=Request.Form("Border_Size")
				    Nav_Links_Seperator=Request.Form("Nav_Links_Seperator")
				    if Request.Form("Nav_Bold") <> "" then
					    Nav_Bold = Request.Form("Nav_Bold")
				    else
					    Nav_Bold = 0
				    end if
				    if Request.Form("Nav_Italic") <> "" then
					    Nav_Italic = Request.Form("Nav_Italic")
				    else
					    Nav_Italic = 0
				    end if
				    call sub_set_defaults()
				    sql_where="Nav_Menu_Number="&Nav_Menu_Number&",Nav_Position='"&Nav_Position&"',Nav_Spacing='"&Nav_Spacing&"',Nav_Font_Face='"&Nav_Font_Face&"',Nav_Font_Size="&Nav_Font_Size&",Nav_Font_Color='"&Nav_Font_Color&"',Color_1='"&Color_1&"',color_2='"&color_2&"',BG='"&BG&"',Border_Size='"&Border_Size&"',Border_Color='"&Border_Color&"',Nav_Links_Seperator='"&Nav_Links_Seperator&"',Nav_Bold="&Nav_Bold&",Nav_Italic="&Nav_Italic&","
			    case "Nav Buttons"
                    Nav_Menu_Number=Request.Form("Nav_Menu_Number")
				    BG=Request.Form("BG")
				    Nav_Buttons_Alignment=Request.Form("Nav_Buttons_Alignment")
				    Nav_Spacing=Request.Form("Nav_Spacing")
				    Object_Width=Request.Form("Object_Width")
				    Object_Height=Request.Form("Object_Height")			
				    Alignment=Request.Form("Alignment")
				    Nav_Position=Request.Form("Nav_Position")
				    Nav_Font_Face=Request.Form("Nav_Font_Face")
				    Nav_Font_Color=Request.Form("Nav_Font_Color")
				    Nav_Font_Size=Request.Form("Nav_Font_Size")
				    Border_Color=Request.Form("Border_Color")
				    Border_Size=Request.Form("Border_Size")
				    Color_1=Request.Form("Color_1")
				    color_2=Request.Form("color_2")
				    if Request.Form("Nav_Bold") <> "" then
					    Nav_Bold = Request.Form("Nav_Bold")
				    else
					    Nav_Bold = 0
				    end if
				    if Request.Form("Nav_Italic") <> "" then
					    Nav_Italic = Request.Form("Nav_Italic")
				    else
					    Nav_Italic = 0
				    end if
				    call sub_set_defaults()
				    sql_where="Nav_Font_Color='"&Nav_Font_Color&"',Nav_Menu_Number="&Nav_Menu_Number&",BG='"&BG&"',Object_Width='"&Object_Width&"', Object_Height='"&Object_Height&"', Alignment='"&Alignment&"', Nav_Position='"&Nav_Position&"',Nav_Spacing="&Nav_Spacing&",Nav_Font_Face='"&Nav_Font_Face&"',Nav_Font_Size="&Nav_Font_Size&",Border_Color='"&Border_Color&"',Border_Size="&Border_Size&", Color_1='"&Color_1&"',color_2='"&color_2&"',Nav_Bold="&Nav_Bold&",Nav_Italic="&Nav_Italic&","
			    case "HTML Text"
				    HTML_Text=nullifyQ(Request.Form("HTML_Text"))
				    call sub_set_defaults()
				    sql_where="HTML_Text='"&HTML_Text&"'," 
			    case "HR"
				    Object_Height=Request.Form("Object_Height")
				    Object_Width=Request.Form("Object_Width")
				    HR_Alignment=Request.Form("HR_Alignment")		
				    Color_1=Request.Form("Color_1")		
				    if Color_1="" then
					    Color_1="000000"
				    end if
				    call sub_set_defaults()
				    sql_where="Object_Height='"&Object_Height&"', Object_Width='"&Object_Width&"', Color_1='"&Color_1&"',"   
			    case else
			        sql_where=""	
			end select

            sql_update="update store_design_objects set "&sql_where&"  Same_line="&Same_Line&", Design_Area='"&Design_Area&"', Object_Area_Height='"&Object_Area_Height&"' , Object_Area_Width='"&Object_Area_Width&"' , Object_Alignment='"&Object_Alignment&"' where Object_Id="&Object_Id&" and Store_Id="&Store_Id&" and template_id="&Template_Id
		  on error goto 0
		    conn_store.execute(sql_update)
		    response.redirect "layout_design.asp?Id="&Template_Id
		else
		    sql_update = "exec wsp_design_object_insert "&store_id&","&template_id&","&design_area&",'"&object_type&"'"
		    conn_store.execute(sql_update)
		    sql_select="select max(object_id) as Object_Id from store_design_objects where Object_Type='"&object_type&"' and Store_Id="&Store_Id&" and template_id="&Template_Id
			rs_Store.open sql_select,conn_store,1,1
	        if not rs_store.eof then
	            Object_Id=rs_store("Object_Id")
	        end if
	        rs_store.close
		    response.redirect "manage_object.asp?Id="&Object_Id&"&op=edit&tid="&Template_Id&"&Object_Type="&object_type	
		end if
	end if	
end if

sub sub_set_defaults ()
	if not isNumeric(Nav_Spacing) then
		Nav_Spacing=0
	end if
	if not isNumeric(Nav_Font_Size) then
		Nav_Font_Size=2
	end if
	if not isNumeric(Border_Size) then
		Border_Size=0
	end if
	if not isNumeric(Nav_Bold) then
		Nav_Bold=0
	end if
	if not isNumeric(Nav_Italic) then
		Nav_Italic=0
	end if
end sub
%>

