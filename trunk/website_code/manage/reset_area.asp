<%
function fn_make_style (Template_Id)
    sql_obj="select * from store_design_objects WITH (NOLOCK) where store_id="&Store_Id&" and Template_Id="&Template_Id&" and (Object_Type='Nav Links' or Object_Type='Nav Buttons') order by view_order"
	
	set rs_obj=conn_store.execute(sql_obj)
	strStyle = ""
	
    while not (rs_obj.EOF)
        Object_Type = rs_Obj("Object_Type")
        if Object_Type= "Nav Links" then	
			    Nav_Links_Alignment=Object_Area_Alignment
			    Nav_Spacing=rs_Obj("Nav_Spacing")
			    BG=rs_Obj("BG")
			    Border_Color=rs_Obj("Border_Color")
			    Nav_Position=rs_Obj("Nav_Position")
			    Nav_Links_Seperator=rs_Obj("Nav_Links_Seperator")
			    Nav_Font_Face=rs_Obj("Nav_Font_Face")
			    Nav_Font_Color=rs_Obj("Nav_Font_Color")
			    Nav_Font_Size=rs_Obj("Nav_Font_Size")
			    Color_1=rs_Obj("Color_1")
			    color_2=rs_Obj("color_2")
			    Nav_Bold=rs_Obj("Nav_Bold")
			    Nav_Italic=rs_Obj("Nav_Italic")
			    Nav_Menu_Number=rs_Obj("Nav_Menu_Number")
			    if Nav_Bold <> 0 then
				    Nav_Bold="font-weight:bold;"
			    else
				    Nav_Bold=""
			    end if
			    if Nav_Italic <> 0 then
				    Nav_Italic="font-style:italic;"
			    else
				    Nav_Italic=""
			    end if
		     
			    if BG <> ""  then
				    if Instr(1, BG, ".") > 0 then
					    BG="background=images/"&rs_Obj("BG")
				    else
					    BG="bgcolor=#"&rs_Obj("BG")
				    end if	
			    else 
				    BG=str_Design_Area_BG
			    end if
			    Border_Color=rs_Obj("Border_Color")
			    Border_Size=rs_Obj("Border_Size")		
		        strStyle=strStyle & vbcrlf&"a.links"& rs_Obj("Object_Id")&_
		            " {color: #"&Nav_Font_Color&_
		            "; font-size:"&Nav_Font_Size&_
		            "pt;font-family :"&Nav_Font_Face&_
		            "; "&Nav_Bold&" "&Nav_Italic&_
		            " text-decoration: none}" &_
		            "a.links"& rs_Obj("Object_Id")&_
		            ":link {color: #"&Nav_Font_Color&_
		            "; text-decoration: none}" &_
		            "a.links"& rs_Obj("Object_Id")&_
		            ":visited {color: #"&Color_1&_
		            "; text-decoration: none}" &_
		            "a.links"& rs_Obj("Object_Id")&_
		            ":hover {color: #"&color_2&_
		            "; text-decoration: underline}" &_
		            " .links"& rs_Obj("Object_Id")&_
		            "box {border-right: #"&Border_Color&" "&Border_Size&_
		            "px solid; border-top: #"&Border_Color&" "&Border_Size&_
		            "px solid; border-left: #"&Border_Color&" "&Border_Size&_
		            "px solid;  border-bottom: #"&Border_Color&" "&Border_Size&_
		            "px solid;} .links"& rs_Obj("Object_Id")&_
		            "sep{color: #"&Nav_Font_Color&_
		            "; font-size:"&Nav_Font_Size&_
		            "pt;font-family :"&Nav_Font_Face&_
		            ";"&Nav_Bold&" "&Nav_Italic&_
		            " text-decoration: none}"&vbcrlf

	        elseif Object_Type= "Nav Buttons" then
			    Border_Color=rs_Obj("Border_Color")
			    Border_Size=rs_Obj("Border_Size")
			    Color_1=rs_Obj("Color_1")
			    color_2=rs_Obj("color_2")
			    Nav_Spacing=rs_obj("Nav_Spacing")
			    Object_Width=rs_obj("Object_Width")
			    Object_Height=rs_obj("Object_Height")
			    Alignment=rs_obj("Alignment")
			    Nav_Position=rs_obj("Nav_Position")
			    BG=rs_Obj("BG")
			    Nav_Font_Face=rs_Obj("Nav_Font_Face")
			    Nav_Bold=rs_Obj("Nav_Bold")
			    Nav_Italic=rs_Obj("Nav_Italic")	
			    Nav_Menu_Number=rs_Obj("Nav_Menu_Number")
			    if Nav_Bold <> 0 then
				    Nav_Bold="font-weight:bold;"
			    else
				    Nav_Bold=""
			    end if
			    if Nav_Italic <> 0 then
				    Nav_Italic="font-style:italic;"
			    else
				    Nav_Italic=""
			    end if

		        strStyle=strStyle & vbcrlf&" .buttons"& rs_Obj("Object_Id")&_
		            " {text-align:"&Alignment&_
		            ";width:"&Object_Width&_
		            ";height:"&Object_Height&_
		            ";background-color: #"&Color_1&_
		            ";color: #"&rs_Obj("nav_font_color")&_
		            "; font-size:"&rs_Obj("Nav_Font_Size")&_
		            "pt;font-family :"&Nav_Font_Face&_
		            "; "&Nav_Bold&" "&Nav_Italic&_
		            " text-decoration: none}" &_
		            ".buttons"& rs_Obj("Object_Id")&_
		            ":link {color: #"&rs_Obj("nav_font_color")&_
		            "; text-decoration: none}" &_
		            ".buttons"& rs_Obj("Object_Id")&_
		            ":hover {background-color: #"&color_2&_
		            ";color: #"&rs_Obj("nav_font_color")&_
		            "; "&Nav_Bold&" "&Nav_Italic&_
		            " text-decoration: underline}" &vbcrlf
        end if
        rs_obj.movenext
    wend
	rs_obj.close
	fn_make_style=strStyle
end function

function fn_make_area (replace_name,Template_Id)
    if replace_name="Top" then
	    design_area=1
	elseif replace_name="Left" then
	    design_area=2
	elseif replace_name="Right" then
	    design_area=3
	elseif replace_name="Bottom" then
	    design_area=4
	elseif replace_name="Cen_Top" then
	    design_area=5
	elseif replace_name="Cen_Bot" then
	    design_area=6
	end if

	sql_obj="select * from store_design_objects where design_area="&design_area&" and store_id="&Store_Id&" and Template_Id="&Template_Id&" order by view_order"
	set rs_obj=conn_store.execute(sql_obj)
	strTop = ""

	if rs_obj.EOF then
		strTop= strTop & "<tr><td valign='top'></td></tr>"
	else
        sameLine=0
		while not (rs_obj.EOF)
            if strTop<>"" and sameLine=1 then
               strTop = strTop & "<td valign='top'>"
            else
                strTop= strTop & "<tr><td valign='top'><table width='100%' border='0' cellspacing='0' cellpadding='0' valign='top'><tr><td valign='top'>"
            end if
			if rs_Obj("Object_Area_Width") <> "0" and rs_Obj("Object_Area_Width") <> "" then
				Object_Area_Width=" width='"&rs_Obj("Object_Area_Width")&"'"
			else
				Object_Area_Width=""
			end if
			if rs_Obj("Object_Area_Height") <> "0" and rs_Obj("Object_Area_Height") <> "" then
				Object_Area_Height=" height='"&rs_Obj("Object_Area_Height")&"'"
			else
				Object_Area_Height="" 
			end if
			str_Object_Area_Alignment=rs_Obj("Object_Alignment")
			if rs_Obj("Object_Alignment") <> "" then
				Object_Area_Alignment=" align="&rs_Obj("Object_Alignment")
			else
				Object_Area_Alignment=""
			end if

		    preObj =vbcrlf&"<!-- obj starts -->"&vbcrlf&"<table name=obj2 cellspacing=0 cellpadding=0 border=0 "&Object_Area_Alignment&" "&Object_Area_Height&" "&Object_Area_Width&" valign='top'>"&vbcrlf&Chr(9)
		    preObj = preObj &"<tr>"&vbcrlf&Chr(9)&Chr(9)
		    preObj = preObj &"	<td valign='top' "&Object_Area_Alignment&">"&vbcrlf&Chr(9)&Chr(9)

		    postObj="</td></tr></table>"&vbcrlf&"<!-- obj ends -->"&vbcrlf

            preimg ="	<table name=obj2 cellspacing=0 cellpadding=0 border=0 height='100%' width='100%' valign='top'>"&vbcrlf&Chr(9)
			preimg = preimg &"<tr>"&vbcrlf&Chr(9)&Chr(9)
			preimg = preimg &"	<td valign='top'>"&vbcrlf&Chr(9)&Chr(9)	
			postimg="</td></tr></table>"																					
			
			pre_HTML="<table "&Object_Area_Width&"  border=0 cellspacing=0 cellpadding=0 "&Object_Area_Height&" valign='top'><tr><td valign=top "&Object_Area_Alignment&">"
			post_HTML="</td></tr></table>"	
							
			select case rs_Obj("Object_Type")
			    case "Banner","Image"
			        sUrl=trim(rs_Obj("Link_URL"))
			        if sUrl<>"" then
					    imgHTML="<a href='"&trim(rs_Obj("Link_Protocol"))&trim(sUrl)&"' target='_"&trim(rs_Obj("Link_Target"))&"'>"
					else
					    imgHTML=""
					end if
					if rs_Obj("Image_Path") = "" then
						imgHTML=imgHTML & rs_Obj("Link_Text")
					else
						sImage_Path=rs_Obj("Image_Path")
						if instr(sImage_Path,"http")=0 then
							sImage_Path="images/"&sImage_Path
						end if
						imgHTML=imgHTML &"<img src='"&sImage_Path&"'"
						if rs_Obj("Alignment") <> "" then
							imgHTML=imgHTML+" align="&rs_Obj("Object_Alignment")
						end if
						if rs_Obj("Image_Alt") <> "" then
							imgHTML=imgHTML+" alt='"&rs_Obj("Image_Alt")&"'"
						end if
						if rs_Obj("Border_Size") <> "" then
							imgHTML=imgHTML+" border='"&rs_Obj("Border_Size")&"'"
						end if
						imgHTML=imgHTML & Object_Area_Height
					    imgHTML=imgHTML & Object_Area_Width
						if rs_Obj("Image_HSpace") <>0 then
							imgHTML=imgHTML+" hspace='"&rs_Obj("Image_HSpace")&"'"
						end if

						if rs_Obj("Image_VSpace") <> 0 then
							imgHTML=imgHTML+" vspace='"&rs_Obj("Image_VSpace")&"'"
						end if
						imgHTML=imgHTML+">"				
					end if
					if sUrl<>"" then					
					    imgHTML=imgHTML+"</a>"
					end if
					strTop=strTop & preObj& imgHTML & postObj
				case "Space"
					str_border=""
					if rs_Obj("Border_Size") <> "" and rs_Obj("Border_Size") <> "0" then
						if rs_Obj("Border_Color") <> "" then
							str_border="style='border:"&rs_Obj("Border_Size")&"px "
						end if
						if rs_Obj("Border_Color") <> "" then
							str_border= str_border +  "solid #"&rs_Obj("Border_Color")
						end if
						str_border=str_border+"'"
					end if					
					spaceHTML="<table cellspacing=0 cellpadding=0 valign='top' "&Object_Area_Width&" "&Object_Area_Height&" "&str_border&" "
					spaceHTML=spaceHTML & "><tr><td valign='top'"
                    swidth=rs_obj("Object_Area_Width")
                    if swidth="" then
                        swidth=0
                    end if
                    sheight=rs_obj("Object_Area_Height")
                    if sheight="" then
                        sheight=0
                    end if
                    BG=rs_Obj("BG")
                    if rs_Obj("BG")="" then
                        BG=""
                    end if
					spaceHTML=spaceHTML & "bgcolor="""&BG&"""><img src='images/images_themes/spacer.gif' height='"&sheight&"' width='"&swidth&"'></td></tr></table>"
					strTop=strTop & preObj& spaceHTML & postObj
				case "HTML Text"
                    strTop=strTop & preObj& rs_Obj("HTML_Text") & postObj
				case "Link"
					linkHTML="<a href='"&trim(rs_Obj("Link_Protocol"))&trim(rs_Obj("Link_URL"))&"' target='_"&trim(rs_Obj("Link_Target"))&"'>"&rs_Obj("Link_Text")&"</a>"
					strTop=strTop & preObj& linkHTML & postObj
				case "Search Box"
					strTop=strTop & preObj&"%OBJ_SEARCH_BOX_OBJ%"& postObj
				case "Login Box"
					strTop=strTop & preObj&"%OBJ_LOGIN_OBJ%"& postObj
				case "Department Select Box"
					strTop=strTop & preObj&"%OBJ_SELECT_BOX_DEPTS_OBJ%"& postObj
				case "Small Cart"
					strTop=strTop & preObj&"%OBJ_SMALL_CART_OBJ%"& postObj
				case "Date"
					strTop=strTop & preObj&"%OBJ_DATE_OBJ%"& postObj
				case "Newsletter Signup"
				    strTop=strTop & preObj&"%OBJ_NEWSLETTER_OBJ%"& postObj
				case "HR"
					if Object_Area_Width = "" then
						Object_Area_Width="width='100%'"
					end if
					hr_HTML="<hr "&Object_Area_Alignment&" "&Object_Area_Width&" "
					if rs_Obj("Object_Area_Height") <> "" and rs_Obj("Object_Area_Height") <> "0" then
						hr_HTML= hr_HTML+" size='"&rs_Obj("Object_Area_Height")&"'" 
					else
						hr_HTML= hr_HTML+" size='1'"
					end if
					if rs_Obj("Color_1") <> "" and rs_Obj("Color_1") <> "0" then
						hr_HTML= hr_HTML+" color='#"&rs_Obj("Color_1")&"'"
					end if
					hr_HTML= hr_HTML+">"
					strTop=strTop & preObj& hr_HTML & postObj
				case "Nav Links"	
				    Nav_Links_Alignment=Object_Area_Alignment
				    Nav_Spacing=rs_Obj("Nav_Spacing")
				    BG=rs_Obj("BG")
				    Border_Color=rs_Obj("Border_Color")
				    Nav_Position=rs_Obj("Nav_Position")
				    Nav_Links_Seperator=rs_Obj("Nav_Links_Seperator")
				    Nav_Font_Face=rs_Obj("Nav_Font_Face")
				    Nav_Font_Color=rs_Obj("Nav_Font_Color")
				    Nav_Font_Size=rs_Obj("Nav_Font_Size")
				    Color_1=rs_Obj("Color_1")
				    color_2=rs_Obj("color_2")
				    Nav_Bold=rs_Obj("Nav_Bold")
				    Nav_Italic=rs_Obj("Nav_Italic")
				    Nav_Menu_Number=rs_Obj("Nav_Menu_Number")
				    if Nav_Bold <> 0 then
					    Nav_Bold="font-weight:bold;"
				    else
					    Nav_Bold=""
				    end if
				    if Nav_Italic <> 0 then
					    Nav_Italic="font-style:italic;"
				    else
					    Nav_Italic=""
				    end if
			        if Design_Area_BG <> ""  then
				        if Instr(1, Design_Area_BG, ".") > 0 then
						        str_Design_Area_BG="background=images/"&Design_Area_BG
				        else
						        str_Design_Area_BG="bgcolor=#"&Design_Area_BG
				        end if	
			        end if
				    if BG <> ""  then
					    if Instr(1, BG, ".") > 0 then
						    BG="background=images/"&rs_Obj("BG")
					    else
						    BG="bgcolor=#"&rs_Obj("BG")
					    end if	
				    else 
					    BG=str_Design_Area_BG
				    end if
				    Border_Color=rs_Obj("Border_Color")
				    Border_Size=rs_Obj("Border_Size")		
			        
				    if trim(Nav_Position) = "Vertical" then
					    pre_linkHTML="<TR><TD valign='top'><table  valign='top' cellspacing="&Nav_Spacing&" cellpadding=0 class=links"& rs_Obj("Object_Id")&"box  "&Nav_Links_Alignment&" "&BG&" "&Object_Area_Width&" "&Object_Area_Height&"><TD valign='top' "&Nav_Links_Alignment&">"
					    Nav_HTML="<TR><TD valign='top' "&Nav_Links_Alignment&"><font class=links"& rs_Obj("Object_Id")&"sep>"&Nav_Links_Seperator&"</font></TD></TR><TR><TD valign='top' "&Nav_Links_Alignment&">%OBJ_LINK_OBJ%</TD></TR>"
					    post_linkHTML="</TD></table></TD></TR>"
				    else				
                        select case str_Object_Area_Alignment 
                            case "Left"
	                            left_empty_buttons=""
	                            right_empty_buttons="<TD valign='top' width='100%' ></TD>"
                            case "Right"
	                            left_empty_buttons="<TD valign='top' width='100%' ></TD>"
	                            right_empty_buttons=""
                            case "Center"
	                            left_empty_buttons="<TD valign='top' width='50%' ></TD>"
	                            right_empty_buttons="<TD valign='top' width='50%' ></TD>"
                        end select
						left_empty_buttons=""
						right_empty_buttons=""
					    pre_linkHTML="<table cellspacing=0 cellpadding=1 class=links"& rs_Obj("Object_Id")&"box "&BG&" "&Object_Area_Width&" "&Object_Area_Height&"><TR>"&left_empty_buttons
						
					    pre_linkHTML=pre_linkHTML+"<TD "&Nav_Links_Alignment&" class=links"& rs_Obj("Object_Id")&"sep>"
					    if Nav_Spacing <> 0 then
				            Nav_HTML= Nav_Links_Seperator&"%OBJ_LINK_OBJ% <img src='images/images_themes/spacer.gif' width="&Nav_Spacing&" height=1>"
					    else
						    Nav_HTML=Nav_Links_Seperator&"%OBJ_LINK_OBJ%"
					    end if
					    post_linkHTML="</TD>"
					    post_linkHTML=post_linkHTML + right_empty_buttons&"</TR></table>"
				    end if			
				    
				    sThisLinks=fn_make_nav ("link",Nav_HTML,rs_Obj("Object_Id"),Nav_Menu_Number)
			
			        strTop=strTop &  preObj &pre_linkHTML & sThisLinks &vbcrlf&post_linkHTML &postObj

		        case "Nav Buttons"	
				    Border_Color=rs_Obj("Border_Color")
				    Border_Size=rs_Obj("Border_Size")
				    Color_1=rs_Obj("Color_1")
				    color_2=rs_Obj("color_2")
				    Nav_Spacing=rs_obj("Nav_Spacing")
				    Object_Width=rs_obj("Object_Width")
				    Object_Height=rs_obj("Object_Height")
				    Alignment=rs_obj("Alignment")
				    Nav_Position=rs_obj("Nav_Position")
				    BG=rs_Obj("BG")
				    Nav_Font_Face=rs_Obj("Nav_Font_Face")
				    Nav_Bold=rs_Obj("Nav_Bold")
				    Nav_Italic=rs_Obj("Nav_Italic")	
				    Nav_Menu_Number=rs_Obj("Nav_Menu_Number")
				    if Nav_Bold <> 0 then
					    Nav_Bold="font-weight:bold;"
				    else
					    Nav_Bold=""
				    end if
				    if Nav_Italic <> 0 then
					    Nav_Italic="font-style:italic;"
				    else
					    Nav_Italic=""
				    end if

			        if Design_Area_BG <> "" then
				        if Instr(1, Design_Area_BG, ".") > 0 then
						        str_Design_Area_BG="background=images/"&Design_Area_BG
				        else
						        str_Design_Area_BG="bgcolor=#"&Design_Area_BG
				        end if	
			        end if
				    if BG <> ""  then
					    if Instr(1, BG, ".") > 0 then

						    BG="background=images/"&BG
					    else
						    BG="bgcolor=#"&BG
					    end if
				    else 
					    BG=str_Design_Area_BG
				    end if
					
					if trim(Nav_Position) = "Horizontal" then
					    select case str_Object_Area_Alignment
						    case "Left"
							    left_empty_buttons=""
						        right_empty_buttons="<TD valign='top' width='100%' height='"&Object_Height&"' style='border:%OBJ_Nav_Buttons_Border_Size_OBJ%px solid #%OBJ_Nav_Buttons_Border_Color_OBJ%'></TD>"
						    case "Right"
							    left_empty_buttons="<TD valign='top' width='100%' height='"&Object_Height&"' style='border:%OBJ_Nav_Buttons_Border_Size_OBJ%px solid #%OBJ_Nav_Buttons_Border_Color_OBJ%'></TD>"
							    right_empty_buttons=""
						    case "Center"
							    left_empty_buttons="<TD valign='top' width='50%' height='"&Object_Height&"' style='border:%OBJ_Nav_Buttons_Border_Size_OBJ%px solid #%OBJ_Nav_Buttons_Border_Color_OBJ%'></TD>"
						        right_empty_buttons="<TD valign='top' width='50%' height='"&Object_Height&"' style='border:%OBJ_Nav_Buttons_Border_Size_OBJ%px solid #%OBJ_Nav_Buttons_Border_Color_OBJ%'></TD>"
					    end select

						Nav_HTML="<TD valign='top'><table valign='top' cellspacing=%OBJ_Nav_Buttons_Border_Size_OBJ% cellpadding=0 width=%OBJ_Nav_Buttons_Width_OBJ% bgcolor=#%OBJ_Nav_Buttons_Border_Color_OBJ% align=center><TR><TD class=buttons"& rs_Obj("Object_Id")&">%OBJ_LINK_OBJ%</TD></TR></table></TD>"
						if Nav_Spacing <> 0 then
						    Nav_HTML=Nav_HTML + "<TD valign='top'><table valign='top' cellspacing=0 cellpadding=0 width="&Nav_Spacing&"><TR><TD valign='top'><img src='images/images_themes/spacer.gif'></TD></TR></table></TD>"
						    left_empty_buttons=left_empty_buttons + "<TD valign='top'><table valign='top' cellspacing=0 cellpadding=0 width="&Nav_Spacing&"><TR><TD valign='top'><img src='images/images_themes/spacer.gif'></TD></TR></table></TD>"
						end if
						prenav="<table valign='top' cellspacing=0 cellpadding=0 border=0 "&Object_Area_Height&" "&Object_Area_Width&" "&BG&"><TR>"&left_empty_buttons&" <TD valign='top' width=50><TABLE cellspacing=0 cellpadding=0 border=0 valign='top'><TR>	"		
						postnav="</TR></TABLE></TD>"&right_empty_buttons&"</TR></table>"	
					else		'Nav Buttons Vertical
						prenav="<table valign='top' cellpadding=0 cellspacing=0 border=0 "&Object_Area_Height&" "&Object_Area_Width&" "&BG&">"
						postnav="</table>"		
						Nav_HTML="<TR><TD valign='top'><table valign='top' cellspacing=%OBJ_Nav_Buttons_Border_Size_OBJ% cellpadding=0 width=%OBJ_Nav_Buttons_Width_OBJ% bgcolor=#%OBJ_Nav_Buttons_Border_Color_OBJ% "&Object_Area_Alignment&"><TR><TD class=buttons"& rs_Obj("Object_Id")&">%OBJ_LINK_OBJ%</TD></TR></table></TD></TR>"

						if Nav_Spacing <> 0 then
							Nav_HTML=Nav_HTML + "<TR><TD height="&Nav_Spacing&"><img src='images/images_themes/spacer.gif'></TD></TR>"
						end if
					end if


				    Nav_HTML =  fn_replace(Nav_HTML,"%OBJ_Nav_Buttons_Border_Size_OBJ%",Border_Size)
				    Nav_HTML =  fn_replace(Nav_HTML,"%OBJ_Nav_Buttons_Border_Color_OBJ%",Border_Color)
				    Nav_HTML =  fn_replace(Nav_HTML,"%OBJ_Nav_Buttons_Alignment_OBJ%",Nav_Buttons_Alignment)
				    Nav_HTML =  fn_replace(Nav_HTML,"%OBJ_Nav_Buttons_Width_OBJ%",Object_Width)
				    Nav_HTML =  fn_replace(Nav_HTML,"%OBJ_Nav_Buttons_Color_OBJ%",Color_1)
				    Nav_HTML =  fn_replace(Nav_HTML,"%OBJ_Nav_Buttons_Active_Color_OBJ%",Color_2)


				    prenav =  Replace(prenav,"%OBJ_Nav_Buttons_Border_Size_OBJ%",Border_Size)
				    prenav =  Replace(prenav,"%OBJ_Nav_Buttons_Border_Color_OBJ%",Border_Color)
				    postnav =  Replace(postnav,"%OBJ_Nav_Buttons_Border_Size_OBJ%",Border_Size)
				    postnav =  Replace(postnav,"%OBJ_Nav_Buttons_Border_Color_OBJ%",Border_Color)
					
					sThisButtons=fn_make_nav ("button",Nav_HTML,rs_Obj("Object_Id"),Nav_Menu_Number)
			        strTop=strTop & preObj & prenav & sThisButtons&vbcrlf&postnav & postObj

			    case else
				    strTop=strTop & "<font size=2 face=Verdana, Arial, Helvetica, sans-serif>"&rs_Obj("Object_Type")&"</font>"
			    end select

			    strTop=strTop &"</td>"
			    sameLine=0
			if rs_Obj("same_line") = 0 then
			    sameLine=0
			    strTop=strTop &vbcrlf &"</tr></table></td></tr>"
			else
                sameLine=1
            end if
			rs_obj.movenext
		Wend
		if sameLine=1 then
		   'if this is the last object in area must close line anyway
		   strTop=strTop &vbcrlf &"</tr></table></td></tr>"
		end if
	end if	'EOF condition
    fn_make_area=strTop
    
end function

function fn_make_nav (sType,sNavTemplate,sObj_Id,sMenu_num)
    
    links = ""
    
    if not isNull(sNavTemplate) then
        sql_select = "wsp_design_menu "&store_id&","&sMenu_num&",'"&stype&"';"
        set myfields=server.createobject("scripting.dictionary")
        Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

	   if sObj_Id="0" and stype="button" then
	   	sClass="nav"
	   elseif sObj_Id="0" and stype="link" then
	   	sClass="link"
	   else
	   	sClass=stype&"s"&sObj_Id
	end if
        if noRecords = 0 then
            FOR rowcounter= 0 TO myfields("rowcount")
                bCreate = 1
                File_name = mydata(myfields("file_name"),rowcounter)
                Page_name = mydata(myfields("page_name"),rowcounter)
                is_link = mydata(myfields("is_link"),rowcounter)
                menu_type = mydata(myfields("menu_type"),rowcounter)
    		
    		    if menu_type="page" then
                    if File_name = "browse_dept_items.asp" then
	                    File_name="items/list.htm"
                    end if

                    File_Name = fn_page_url (File_Name,is_link)
                else
                    File_Name = fn_dept_url(File_Name,"")
                end if
                if bCreate then
	                sLinkText = vbcrlf&"<a href='"&File_name&"' class='"&sClass&"'>"&Page_name&"</a>"
	                links = links & Replace(sNavTemplate,"%OBJ_LINK_OBJ%",sLinkText)
			    end if
            Next
        End if
    end if
    
    fn_make_nav=links

end function

%>
