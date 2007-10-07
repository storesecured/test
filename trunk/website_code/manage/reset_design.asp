<!--#include file="global_settings.asp"-->
<!--#include file="reset_area.asp"-->

<%
if Session("Preview")<>"" then
	sPreview=Session("Preview")
else
	sPreview=0
end if

SQL_store_settings = "exec wsp_design_select "&Store_Id&" , "&sPreview
rs_Store.open SQL_store_settings,conn_store,1,1
if rs_store.eof then
	'if were here then no store was found so exit
	rs_store.close
else
    Template_Id = rs_Store("Template_Id")
    Global_Bck_Color = rs_Store("Global_Bck_Color")
    Template_Html  = rs_Store("Template_Html")
    Template_Head  = rs_Store("Template_Head")
    
    Global_Text_Font_Size  = rs_Store("Global_Text_Font_Size")
    Global_Text_Font_Color  = rs_Store("Global_Text_Font_Color")
    Global_Text_Font_Face  = rs_Store("Global_Text_Font_Face")
    Global_Link_Font_Size  = rs_Store("Global_Link_Font_Size")
    Global_Link_Font_Color  = rs_Store("Global_Link_Font_Color")
    Global_Link_Font_Face  = rs_Store("Global_Link_Font_Face")
    Global_Nav_Font_Size  = rs_Store("Global_Nav_Font_Size")
    Global_Nav_Font_Color  = rs_Store("Global_Nav_Font_Color")
    Global_Nav_Font_Face = rs_Store("Global_Nav_Font_Face")
    
    Design_Top_BG = rs_store("Design_Top_BG")
	Design_Left_BG = rs_store("Design_Left_BG")
	Design_Right_BG = rs_store("Design_Right_BG")
	Design_Bottom_BG = rs_store("Design_Bottom_BG")
	Design_Cen_Top_BG = rs_store("Design_Cen_Top_BG")
	Design_Cen_Bot_BG = rs_store("Design_Cen_Bot_BG")
	Design_Cen_Cen_BG = rs_store("Design_Cen_Cen_BG")
	
	Design_Top_Border=rs_store("Design_Top_Border")
	Design_Top_Border_Color=rs_store("Design_Top_Border_Color")
	Design_Top_Height=rs_store("Design_Top_Height")
	Design_Top_Width=rs_store("Design_Top_Width")
	
	Design_Left_Border=rs_store("Design_Left_Border")
	Design_Left_Border_Color=rs_store("Design_Left_Border_Color")
	Design_Left_Height=rs_store("Design_Left_Height")
	Design_Left_Width=rs_store("Design_Left_Width")

	Design_Cen_Top_Border=rs_store("Design_Cen_Top_Border")
	Design_Cen_Top_Border_Color=rs_store("Design_Cen_Top_Border_Color")
	Design_Cen_Top_Height=rs_store("Design_Cen_Top_Height")
	Design_Cen_Top_Width=rs_store("Design_Cen_Top_Width")

	Design_Cen_Cen_Border=rs_store("Design_Cen_Cen_Border")
	Design_Cen_Cen_Border_Color=rs_store("Design_Cen_Cen_Border_Color")
	Design_Cen_Cen_Height=rs_store("Design_Cen_Cen_Height")
	Design_Cen_Cen_Width=rs_store("Design_Cen_Cen_Width")
	
	Design_Cen_Bot_Border=rs_store("Design_Cen_Bot_Border")
	Design_Cen_Bot_Border_Color=rs_store("Design_Cen_Bot_Border_Color")
	Design_Cen_Bot_Height=rs_store("Design_Cen_Bot_Height")
	Design_Cen_Bot_Width=rs_store("Design_Cen_Bot_Width")

	Design_Right_Border=rs_store("Design_Right_Border")
	Design_Right_Border_Color=rs_store("Design_Right_Border_Color")
	Design_Right_Height=rs_store("Design_Right_Height")
	Design_Right_Width=rs_store("Design_Right_Width")
	
	Design_Bottom_Border=rs_store("Design_Bottom_Border")
	Design_Bottom_Border_Color=rs_store("Design_Bottom_Border_Color")
	Design_Bottom_Height=rs_store("Design_Bottom_Height")
	Design_Bottom_Width=rs_store("Design_Bottom_Width")
	
	nav_link_html=rs_store("nav_link_html")
    nav_button_html=rs_store("nav_button_html")
    rs_store.close

    sDefaultTextFontSize = Global_Text_Font_Size*5
    sDefaultLinkFontSize = Global_Link_Font_Size*5
    sDefaultNavFontSize = Global_Nav_Font_Size*5

    sStyle = vbcrlf&_
	    ".normal {COLOR: #"&Global_Text_Font_Color&_
	    ";TEXT-DECORATION: none;font-size : "&sDefaultTextFontSize&_
	    "pt;font-family :"&Global_Text_Font_Face&_
	    ";}td {COLOR: #"&Global_Text_Font_Color&_
	    ";TEXT-DECORATION: none;font-size : "&sDefaultTextFontSize&_
	    "pt;font-family :"&Global_Text_Font_Face&_
	    ";}.big {COLOR: #"&Global_Text_Font_Color&_
	    ";TEXT-DECORATION: none;font-size : "&sDefaultTextFontSize+2&_
	    "pt;font-family :"&Global_Text_Font_Face&_
	    ";}.small {COLOR: #"&Global_Text_Font_Color&_
	    ";TEXT-DECORATION: none;font-size : "&sDefaultTextFontSize-2&_
	    "pt;font-family :"&Global_Text_Font_Face&_
	    ";}.link {COLOR: #"&Global_Link_Font_Color&_
	    ";TEXT-DECORATION: underline;font-size : "&sDefaultLinkFontSize&_
	    "pt;font-family :"&Global_Link_Font_Face&_
	    ";}.nav {COLOR: #"&Global_Nav_Font_Color&_
	    ";TEXT-DECORATION: none;font-size : "&sDefaultNavFontSize&_
	    "pt;font-family :"&Global_Nav_Font_Face&";}"&_
	    fn_make_style(Template_Id)&vbcrlf
    
    sStyleTotal = "<style type=""text/css"">"&sStyle&"</style>"&vbcrlf

    if Instr(1, Global_Bck_Color, ".") > 0 then
        sGlobalBG = " background=""images/"&Global_Bck_Color&""" "
    		sGlobalBGImage="images/"&Global_Bck_Color
          sGlobalBGColor=""
    else
        sGlobalBG = " bgcolor=""#"&Global_Bck_Color&""" "
        sGlobalBGColor ="#"&Global_Bck_Color
        sGlobalBGImage=""
    end if
    
    sHead=Template_Head&"%OBJ_END_HEAD_OBJ%"&vbcrlf&vbcrlf&Template_Html&"</body></html>"
    if not isNull(sHead) then
        sHead = fn_replace (sHead,"%OBJ_GLOBAL_STYLE_OBJ%",sStyleTotal)
        sHead = fn_replace (sHead,"%OBJ_STYLE_OBJ%",sStyle)
        sHead = fn_replace (sHead,"%OBJ_GLOBAL_BG_OBJ%",sGlobalBG)
        sHead = fn_replace (sHead,"%OBJ_BG_COLOR_OBJ%",sGlobalBGColor)
        sHead = fn_replace (sHead,"%OBJ_BG_IMAGE_OBJ%",sGlobalBGImage)
        sHead = fn_replace (sHead,"OBJ_HEAD_OBJ","")
	   sHead = fn_replace (sHead,"</head>","OBJ_HEAD_OBJ"&vbcrlf&"</head>")

        sHead = fn_replace_background ("%OBJ_TOP_BG%",Design_Top_BG,sHead)
        sHead = fn_replace_background ("%OBJ_CENTOP_BG%",Design_Cen_Top_BG,sHead)
        sHead = fn_replace_background ("%OBJ_LEFT_BG%",Design_Left_BG,sHead)
        sHead = fn_replace_background ("%OBJ_RIGHT_BG%",Design_Right_BG,sHead)
        sHead = fn_replace_background ("%OBJ_CENBOT_BG%",Design_Cen_Bot_BG,sHead)
        sHead = fn_replace_background ("%OBJ_BOTTOM_BG%",Design_Bottom_BG,sHead)
        sHead = fn_replace_background ("%OBJ_CENCEN_BG%",Design_Cen_Cen_BG,sHead)
        sHead = fn_replace(sHead,"%OBJ_CONTENT_WIDTH_OBJ%",Content_Width)

        sHead = fn_replace_border ("%OBJ_TOP_BORDER%",Design_Top_Border,Design_Top_Border_Color,sHead)
        sHead = fn_replace_border ("%OBJ_LEFT_BORDER%",Design_Left_Border,Design_Left_Border_Color,sHead)
        sHead = fn_replace_border ("%OBJ_CENTOP_BORDER%",Design_Cen_Top_Border,Design_Cen_Top_Border_Color,sHead)
        sHead = fn_replace_border ("%OBJ_CENCEN_BORDER%",Design_Cen_Cen_Border,Design_Cen_Cen_Border_Color,sHead)
        sHead = fn_replace_border ("%OBJ_CENBOT_BORDER%",Design_Cen_Bot_Border,Design_Cen_Bot_Border_Color,sHead)
        sHead = fn_replace_border ("%OBJ_RIGHT_BORDER%",Design_Right_Border,Design_Right_Border_Color,sHead)
        sHead = fn_replace_border ("%OBJ_BOTTOM_BORDER%",Design_Bottom_Border,Design_Bottom_Border_Color,sHead)

        sHead = fn_replace_height ("%OBJ_TOP_HEIGHT%",Design_Top_Height,sHead)
        sHead = fn_replace_height ("%OBJ_LEFT_HEIGHT%",Design_Left_Height,sHead)
        sHead = fn_replace_height ("%OBJ_CENTOP_HEIGHT%",Design_Cen_Top_Height,sHead)
        sHead = fn_replace_height ("%OBJ_CENCEN_HEIGHT%",Design_Cen_Cen_Height,sHead)
        sHead = fn_replace_height ("%OBJ_CENBOT_HEIGHT%",Design_Cen_Bot_Height,sHead)
        sHead = fn_replace_height ("%OBJ_RIGHT_HEIGHT%",Design_Right_Height,sHead)
        sHead = fn_replace_height ("%OBJ_BOTTOM_HEIGHT%",Design_Bottom_Height,sHead)

        sHead = fn_replace_width ("%OBJ_TOP_WIDTH%",Design_Top_Width,sHead)
        sHead = fn_replace_width ("%OBJ_LEFT_WIDTH%",Design_Left_Width,sHead)
        sHead = fn_replace_width ("%OBJ_CENTOP_WIDTH%",Design_Cen_Top_Width,sHead)
        sHead = fn_replace_width ("%OBJ_CENCEN_WIDTH%",Design_Cen_Cen_Width,sHead)
        sHead = fn_replace_width ("%OBJ_CENBOT_WIDTH%",Design_Cen_Bot_Width,sHead)
        sHead = fn_replace_width ("%OBJ_RIGHT_WIDTH%",Design_Right_Width,sHead)
        sHead = fn_replace_width ("%OBJ_BOTTOM_WIDTH%",Design_Bottom_Width,sHead)

        sHead = fn_replace_area ("%OBJ_TOP_DES_OBJECTS_OBJ%","Top",sHead,Template_Id)
        sHead = fn_replace_area ("%OBJ_LEFT_DES_OBJECTS_OBJ%","Left",sHead,Template_Id)
        sHead = fn_replace_area ("%OBJ_CENTOP_DES_OBJECTS_OBJ%","Cen_Top",sHead,Template_Id)
        sHead = fn_replace_area ("%OBJ_CENBOT_DES_OBJECTS_OBJ%","Cen_Bot",sHead,Template_Id)
        sHead = fn_replace_area ("%OBJ_RIGHT_DES_OBJECTS_OBJ%","Right",sHead,Template_Id)
        sHead = fn_replace_area ("%OBJ_BOTTOM_DES_OBJECTS_OBJ%","Bottom",sHead,Template_Id)


        sHead=image_replace(sHead)
        sHead = nullifyQ(sHead)
    else
        sHead=""
    end if
    
    template_navbutton = fn_make_nav ("button",nav_button_html,0,1)
    template_navlink = fn_make_nav ("link",nav_link_html,0,1)

    template_navbutton=nullifyQ(template_navbutton)
    template_navlink=nullifyQ(template_navlink)

    if sPreview = "0" then
        sPreview=0
    else
        sPreview=1
    end if

    cPos = InStr(sHead, "%OBJ_END_HEAD_OBJ%")
    sHeadOnly = Left(sHead,cPos-1)&"%OBJ_PAGE_TOP_OBJ%%OBJ_CENTER_CONTENT_OBJ%%OBJ_PAGE_FORM_OBJ%%OBJ_PAGE_BOTTOM_OBJ%</body></html>"
    sHead = replace(sHead,"%OBJ_END_HEAD_OBJ%","")


    on error goto 0
    sql_update = "exec wsp_template_update "&Store_id&","&sPreview&",'"&sHead&"','"&sHeadOnly&"','"&template_navbutton&"','"&template_navlink&"';"
    conn_store.Execute sql_update
    Session("Preview") = ""
end if

function fn_replace_area (sReplace,sArea,sOrigText,Template_Id)
    if instr(sHead,sReplace)>0 then
    	   sMyArea = fn_make_area(sArea,Template_Id)
        sOrigText = fn_replace(sOrigText, sReplace, sMyArea)
        if sMyArea="<tr><td valign='top'></td></tr>" then
        	if sArea="Cen_Top" or sArea="Cen_Bot" then
        		'if nothing is in center top or bottom area comment it out to remove spacing 
        		sOrigText=fn_replace(sOrigText,"<"&"!-- "&sArea&" starts --"&">","<"&"!-- "&sArea&" starts")
        		sOrigText=fn_replace(sOrigText,"<"&"!-- "&sArea&" ends --"&">",sArea&" ends --"&">")
        	end if
        end if
    end if
    fn_replace_area=sOrigText
end function

function fn_replace_background (sReplace,sBackground,sOrigText)
    sBackground=trim(sBackground)
    if sBackground = "" then
        sOrigText=replace(sOrigText, sReplace,"")
    elseif Instr(1, sBackground, ".") > 0 then
        sOrigText=replace(sOrigText, sReplace,"background='images/"&sBackground&"'")
	else
		sOrigText=replace(sOrigText, sReplace,"bgcolor='#"&sBackground&"'")
    end if
    fn_replace_background=sOrigText
end function

function fn_replace_height (sReplace,sHeight,sOrigText)
    if sHeight <> "0" and sHeight <> ""  then
	    sOrigText=replace(sOrigText, sReplace,"height='"&sHeight&"'")
    else
	    sOrigText=replace(sOrigText, sReplace,"") 
    end if
    fn_replace_height=sOrigText 
end function

function fn_replace_width (sReplace,sWidth,sOrigText)
    if sWidth <> "0" and sWidth <> ""  then
	    sOrigText=replace(sOrigText, sReplace,"width='"&sWidth&"'")
    else
	    sOrigText=replace(sOrigText, sReplace,"width='100%'")
    end if
    fn_replace_width=sOrigText 
end function

function fn_replace_border (sReplace,sBorder,sBorderColor,sOrigText)
    if sBorder <> "" then
        sOrigText=replace(sOrigText, sReplace,"style=""border:"&sBorder&"px solid #"&sBorderColor&";""")
    else	
	    sOrigText=replace(sOrigText, sReplace,"")
    end if
    fn_replace_border=sOrigText
end function

%>
