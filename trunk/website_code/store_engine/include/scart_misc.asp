<%
function fn_create_dept (sTemplateStart,q_Dept_Page_Name,sql_statement,Hide_Empty_Depts,sUrl,Dept_rows,dept_display,number_of_recordset,Dept_Start_Row,Dept_End_Row)
    set deptfields4=server.createobject("scripting.dictionary")
    fn_print_debug "fn_create_dept = "&sql_statement
    Call DataGetRows(conn_store,sql_statement,deptdata4,deptfields4,noRecords4)   
    number_of_recordset = deptfields4("rowcount")+1

    Dept_Start_Row = fn_get_start_row(number_of_recordset,dept_display+1,iDeptRowPerPage,"")
    Dept_End_Row = fn_get_end_row(number_of_recordset,dept_display+1,iDeptRowPerPage,Dept_Start_Row,"")

    paging_dept = fn_paging (number_of_recordset,Dept_Start_Row,Dept_End_Row)
    
    Jump_To=fn_get_querystring("Jump_To")
    if paging_dept then
        sDeptNavigation = fn_make_navigation(sThisUrl,number_of_recordset,dept_display+1,iDeptRowPerPage,Dept_Start_Row,Dept_End_Row,"")
        sDeptJumpItems = fn_jump_menu1(sThisUrl,Jump_To)
        sDeptPage_List = fn_make_pagelist (sFirstPage,sPagePattern,number_of_recordset,dept_display+1,iDeptRowPerPage)
    elseif Jump_To<>"" then
        sDeptJumpItems = fn_jump_menu1(sThisUrl,Jump_To)
    end if
    
    sDeptText=sDeptNavigation & sDeptJumpItems
    	
    ' LOOP showing the Sub-Departments
    ' -------------------------------------------
    Relative_Row =Dept_Start_Row
    if noRecords4=0 then
        sDeptText = sDeptText & "<table border='0' width='100%' >"
	    deptrowcounter4=Dept_Start_Row
	    i=0
	    Dept_rows = (Dept_End_Row-Dept_Start_Row)/(dept_display+1)
	    
	    Do while i<Dept_rows
	        j=1
		    rowCount="Odd"
		    if j<=dept_display+1 and deptrowcounter4<number_of_recordset then
			    sDeptText = sDeptText & "<tr valign=top>"  
			    do while j<=dept_display+1
			        if deptrowcounter4>deptfields4("rowcount") then
					    fn_print_debug "in exit do"
					    Exit Do
				    end if
				    
                    full_name=deptdata4(deptfields4("full_name"),deptrowcounter4)
					image_Path=deptdata4(deptfields4("department_image_path"),deptrowcounter4)
					Department_Name = deptdata4(deptfields4("department_name"),deptrowcounter4)
				    Department_Description = deptdata4(deptfields4("department_description"),deptrowcounter4)
				    last_level = deptdata4(deptfields4("last_level"),deptrowcounter4)
				    sDepartment_Id = deptdata4(deptfields4("department_id"),deptrowcounter4)
				        
				    if image_Path <> "" Then
					    if Instr(image_Path,"http://") > 0 then
						    image = "<IMG Src='"&Image_Path&"' border=0>"
					    else
						    image = "<IMG Src='"&Switch_Name&"images/"&image_path&"' border=0>"
					    end if
				    else
					    image = ""
				    end if
				    
				    sTemplate = sTemplateStart

                    if Hide_Empty_Depts or instr(sTemplate,"OBJ_DEPT_STATUS_OBJ")>0 or instr(sTemplate,"OBJ_DEPT_COUNT_OBJ")>0 then
                        if last_level = 0 then
						    status = ">>"
						    sCount=""
					    else
						    sCount = fn_get_item_count(sDepartment_Id)
						    if sCount=0 then
							        status="(No Items)"
						    else
							        status="(View Items)"
						    end if
                        end if
                    else
                        status="(View Items)"
                    end if
					if status = "(No Items)" and Hide_Empty_Depts then
						fn_print_debug "in no items"
					else
						Department_Link = "<a href='"&fn_dept_url(full_name,"")&"' class='link'>"
						Department_Name_Link = Department_Link & Department_Name & "</a>"
						if image <> "" then
							image = Department_Link & image & "</a>"
						end if
						
				        sWidth = cint(100/(dept_display + 1))
				        sWidth = sWidth & "%"
				        
						sTemplate = fn_replace(sTemplate,"OBJ_IMAGE_OBJ",image)
						sTemplate = fn_replace(sTemplate,"OBJ_DEPT_NAME_OBJ",Department_Name_Link)
						fn_print_debug "Department_Name_Link="&Department_Name_Link
						sTemplate = fn_replace(sTemplate,"OBJ_DEPT_DESC_OBJ",Department_Description)
						sTemplate = fn_replace(sTemplate,"OBJ_DEPT_COUNT_OBJ",sCount)
						sTemplate = fn_replace(sTemplate,"OBJ_DEPT_STATUS_OBJ",status)
						sDeptText = sDeptText & "<TD width='"&sWidth&"'>"&sTemplate&"</td>"
 
					    'RELATE TO NAVIGATIONAL BAR
					    Relative_Row = Relative_Row + 1
					    If Relative_Row >= Dept_End_Row then
						    Exit Do
					    End if

					    j=j+1
					end if
					
					deptrowcounter4=deptrowcounter4+1    
					
                    if response.isclientconnected=false then
                       response.end 
                    end if
				loop 
			end if
			i=i+1
		loop
	    sDeptText = sDeptText & "</table>"
    end if
	
	sDeptText=sDeptText & sDeptNavigation & sDeptJumpItems 

    set deptfields4 = Nothing
    
    fn_create_dept = sDeptText
	
end function

function fn_create_item (sTemplateStart,sql_statement,bSearch,sUrl,iItemRowPerPage,item_display,sMakeNav)

    sItemText=""
   
    if bSearch then
    	sTemplateStart = fn_replace(sTemplateStart,"OBJ_ACCESSORY_OBJ","")
	    sTemplateStart = fn_replace(sTemplateStart,"OBJ_ATTRIBUTE","<!--OBJ_ATTRIBUTE-->")
	    sTemplateStart = fn_replace(sTemplateStart,"OBJ_ORDER_OBJ","")
	    sTemplateStart = fn_replace(sTemplateStart,"OBJ_QTY_OBJ","")
	    sTemplateStart = fn_replace(sTemplateStart,"OBJ_UD_OBJ","")
    end if

    if show_name then
        sItemText = sItemText & "<table border='0' width='100%' cellspacing=0 cellpadding=0><tr>"&_
		    "<td colspan='7' height='16' class='big'><b>"&Sub_Department_name&sExtraName&"</b></td></tr></table>"
    end if
    Jump_To_Items=fn_get_querystring("Jump_To_Items")
	
	fn_print_debug sql_statement
	set itemfields=server.createobject("scripting.dictionary")
	Call DataGetRows(conn_store,sql_statement,itemdata,itemfields,noRecords)
     if err.number<>0 then
     	fn_print_debug "error="&err.description	
     end if
     number_of_recordset_items = itemfields("rowcount")+1

	if noRecords = 1 and Jump_To_Items="" then
	    'there are no items to create
		fn_print_debug "there are no items, exiting function"
		if CurrentFileName = "search_items.asp" then
	        sItemText = sItemText & "No items were found matching your search."
	        sItemText = sItemText &  "<BR><BR><a href=search.asp class=link>Search Again</a>"
        else
			sItemText = ""
		end if
		fn_create_item = sItemText
		exit function
	else
	    number_of_recordset_items = itemfields("rowcount")+1
	    Start_Row_Items = fn_get_start_row(number_of_recordset_items,item_display+1,iItemRowPerPage,"_Items")
        End_Row_Items = fn_get_end_row(number_of_recordset_items,item_display+1,iItemRowPerPage,Start_Row_Items,"_Items")
                        
        if sMakeNav then
            paging_items = fn_paging (number_of_recordset_items,Start_Row_Items,End_Row_Items)
        else
            paging_items=0
        end if
        if paging_items then
	        sNavigation = fn_make_navigation_item(sFirstPage,number_of_recordset_items,item_display+1,iItemRowPerPage,Start_Row_Items,End_Row_Items,"_Items")
	        sJumpItems = fn_jump_menu1(sFirstPage,Jump_To_Items)
	        sPage_List = fn_make_pagelist (sFirstPage,sPagePattern,number_of_recordset_items,item_display+1,iItemRowPerPage)
        elseif Jump_To_Items<>"" then
            sJumpItems = fn_jump_menu1(sThisUrl,Jump_To_Items)
        end if

        sItemText = sItemText & sNavigation
        sItemText = sItemText & sJumpItems
		itemrowcounter=Start_Row_Items
		if isNull(End_Row_Items) then
			End_Row_Items = number_of_recordset_items
		end if
		Relative_Row=Start_Row_Items
		Total_Rows=End_Row_Items-Start_Row_Items
		rows = (End_Row_Items-Start_Row_Items)/(item_display+1)
		sItemText = sItemText & "<table border=0 width='100%'>"
		i=0

		Do while  i<= rows
		    sItemText = sItemText & "<tr valign=top>"
			j=1
			do while j<=item_display+1 and cint(itemrowcounter) < cint(End_Row_Items)
				oldi=i
				if (cint(itemrowcounter)>cint(End_Row_Items+1)) or (cint(itemrowcounter)>number_of_recordset_items) then
					Exit Do	
				end if
				sTemplate = sTemplateStart
				Item_Id = itemdata(itemfields("item_id"),itemrowcounter)
		        sDeptId = itemdata(itemfields("sub_department_id"),itemrowcounter)
		        sName = itemdata(itemfields("item_name"),itemrowcounter)
		        Full_name = itemdata(itemfields("full_name"),itemrowcounter)
		        if instr(Full_Name," > ")>0 then
		        	FullNameArray=split(Full_name," > ")
				sDeptName=FullNameArray(ubound(FullNameArray))
				sDeptName=Full_Name
		        else
			   	sDeptName=Full_Name
			   end if

		        item_page_name=itemdata(itemfields("item_page_name"),itemrowcounter)
		        imagel_path=itemdata(itemfields("imagel_path"),itemrowcounter)
		        images_path=itemdata(itemfields("images_path"),itemrowcounter)
		        if isNull(imagel_path) or imagel_path = "" then
			        imagel_path = images_path
		        end if
		        sItem_Sku=itemdata(itemfields("item_sku"),itemrowcounter)
		        sItemName = itemdata(itemfields("item_name"),itemrowcounter)
		        sItem_Weight = itemdata(itemfields("item_weight"),itemrowcounter)
		        sHide_Price = itemdata(itemfields("hide_price"),itemrowcounter)
		        sEnable_Cust_Price = itemdata(itemfields("enable_cust_price"),itemrowcounter)
		        sUse_Price_By_Matrix = itemdata(itemfields("use_price_by_matrix"),itemrowcounter)
		        sQuantityMin = itemdata(itemfields("quantity_minimum"),itemrowcounter)
		        sConfigs_Exist = itemdata(itemfields("configs_exist"),itemrowcounter)
		        sAttribs_Exist = itemdata(itemfields("attributes_exist"),itemrowcounter)
		        retail_price = itemdata(itemfields("retail_price"),itemrowcounter)
		        sQuantity_Control = itemdata(itemfields("quantity_control"),itemrowcounter)
		        sHide_Stock_This = itemdata(itemfields("hide_stock"),itemrowcounter)
		        exStock= itemdata(itemfields("quantity_in_stock"),itemrowcounter)
		        sQuantity_InStock= itemdata(itemfields("quantity_in_stock"),itemrowcounter)
		        sQuantity_Control_Number = itemdata(itemfields("quantity_control_number"),itemrowcounter)
		        M_d_1 = itemdata(itemfields("m_d_1"),itemrowcounter)
		        M_d_2 = itemdata(itemfields("m_d_2"),itemrowcounter)
		        M_d_3 = itemdata(itemfields("m_d_3"),itemrowcounter)
		        M_d_4 = itemdata(itemfields("m_d_4"),itemrowcounter)
		        M_d_5 = itemdata(itemfields("m_d_5"),itemrowcounter)
		        u_d_1_name = itemdata(itemfields("u_d_1_name"),itemrowcounter)
		        u_d_2_name = itemdata(itemfields("u_d_2_name"),itemrowcounter)
		        u_d_3_name = itemdata(itemfields("u_d_3_name"),itemrowcounter)
		        u_d_4_name = itemdata(itemfields("u_d_4_name"),itemrowcounter)
		        u_d_5_name = itemdata(itemfields("u_d_5_name"),itemrowcounter)
		        use_u_d_1 = itemdata(itemfields("u_d_1"),itemrowcounter)
		        use_u_d_2 = itemdata(itemfields("u_d_2"),itemrowcounter)
		        use_u_d_3 = itemdata(itemfields("u_d_3"),itemrowcounter)
		        use_u_d_4 = itemdata(itemfields("u_d_4"),itemrowcounter)
		        use_u_d_5 = itemdata(itemfields("u_d_5"),itemrowcounter)
		        sDescriptionL = itemdata(itemfields("description_l"),itemrowcounter)
		        sDescriptionS = itemdata(itemfields("description_s"),itemrowcounter)
		        if isNull(sDescriptionL) or sDescriptionL = "" then
			        sDescriptionL = sDescriptionS
		        end if
		        accessories_exist = itemdata(itemfields("accessories_exist"),itemrowcounter)
		        sItem_Handling = itemdata(itemfields("item_handling"),itemrowcounter)
		        sRecurring_Fee= itemdata(itemfields("recurring_fee"),itemrowcounter)
		        sRecurring_Days= itemdata(itemfields("recurring_days"),itemrowcounter)
		        sSpecial_Price_Discount = itemdata(itemfields("retail_price_special_discount"),itemrowcounter)
		        sSpecial_Start_Date = itemdata(itemfields("special_start_date"),itemrowcounter)
		        sSpecial_End_Date = itemdata(itemfields("special_end_date"),itemrowcounter)
		        sFractional = itemdata(itemfields("fractional"),itemrowcounter)
		        sSpecialPrice = itemdata(itemfields("item_price"),itemrowcounter)
		        sDetailUrl = fn_item_url(Full_name,item_page_name)
                  sDeptUrl = fn_dept_url(Full_name,"")

		        stock_corection = -1
		        if instr(sTemplate,"OBJ_IMAGES_OBJ") then
				   sImage = fn_create_image(images_Path,sItemName)
				   if sImage<>"" and sImage<>"&nbsp;" then
				   	sImage="<a href='"&sDetailUrl&"'>"&sImage&"</a>"
				   end if
		             sTemplate = fn_replace(sTemplate,"OBJ_IMAGES_OBJ",sImage)
			   end if
			   if instr(sTemplate,"OBJ_IMAGES_URL_OBJ") then
				   sImage = images_Path
				   if sImage<>"" then
				   	sImage=Switch_Name&"images/"&sImage
				   end if
				   sImage = Switch_Name&"images/"&images_Path

		             sTemplate = fn_replace(sTemplate,"OBJ_IMAGES_URL_OBJ",sImage)
			   end if
		        if instr(sTemplate,"OBJ_IMAGEL_OBJ") then
		        	sTemplate = fn_replace(sTemplate,"OBJ_IMAGEL_OBJ",fn_create_image(imagel_Path,sItemName))
		        end if
		        if instr(sTemplate,"OBJ_IMAGEL_URL_OBJ") then
				   if imagel_Path<>"" then
				   	sImage = imagel_Path
				   else
				     sImage = images_Path
				   end if
				   if sImage<>"" then
				   	sImage=Switch_Name&"images/"&sImage
				   end if
		             sTemplate = fn_replace(sTemplate,"OBJ_IMAGEL_URL_OBJ",sImage)
			   end if
		        sTemplate = fn_replace(sTemplate,"OBJ_SKU_OBJ",sItem_Sku)
		        sTemplate = fn_replace(sTemplate,"OBJ_ID_OBJ",Item_Id)
		        sTemplate = fn_replace(sTemplate,"OBJ_NAME_NOLINK_OBJ",sItemName)
		        sTemplate = fn_replace(sTemplate,"OBJ_NAME_OBJ","<a href='"&sDetailUrl&"' class='link'>"&sName&"</a><br>")
		        sTemplate = fn_replace(sTemplate,"OBJ_DETAIL_LINK_OBJ",sDetailUrl)
                  sTemplate = fn_replace(sTemplate,"OBJ_DEPT_LINK_OBJ",sDeptUrl)
                  sTemplate = fn_replace(sTemplate,"OBJ_DEPT_NAME_OBJ",sDeptName)

		        sHiddenFields = "<input type='Hidden' name='Redirect_to_cart' value='"&Redirect_to_cart&"'>"&_
			          "<Input type='hidden' name='Item_Id' value='"&Item_Id&"'>"&_
			          "<Input type='hidden' name='Item_Sku' value='"&sItem_Sku&"'>"&_
			          "<Input type='hidden' name='Item_Name' value='"&sItemName&"'>"&_
			          "<Input type='hidden' name='Item_Weight' value='"&sItem_Weight&"'>"&_
			          "<a name='"&Item_Id&"'>"
		        if "http://"&Request.ServerVariables("HTTP_HOST")&"/"=Site_Name then
			        'cookie check wont work if posting to different domain
			        sHiddenFields = sHiddenFields & "<Input type='hidden' name='Shopper_Id' value='"&Shopper_Id&"'>"
		        end if
		        sForm = "<form method='POST' action='"&Switch_Name&"cart_action.asp' name='ITEM_"&Item_Id&"'>"
		        sFormName = "ITEM_"&Item_Id
		        if not sHide_Price or isNull(sHide_Price) then
			        sOrderButton = fn_create_action_button ("Button_image_Order", "Insert_To_Cart", "Add To Cart")
			        sQuantityInput = "<INPUT type='text' name='quantity' value='"&sQuantityMin&"' size='2' onKeyPress=""return goodchars(event,'0123456789.')"">"
			        sQuantityHidden = "<INPUT type=hidden name='quantity' value='"&sQuantityMin&"' size='2'>"
		        else
			   	sOrderButton = ""
                    sQuantityHidden = "<INPUT type=hidden name='quantity' value='"&sQuantityMin&"' size='2'>"
				sQuantityInput = sQuantityHidden
			   end if

                  aprice_dif = 0
                  aprice_dif_retail = 0
		        if instr(sTemplate,"OBJ_ATTRIBUTE") and sAttribs_Exist<>0 then
			        if Service_Type => 5 then
				        theFileName = sUrl
				        response.write vbcrlf 
					   %><!--#include file="item_attr_js.asp"-->
				        <!--#include file="item_attr_disp.asp"--><%
			        end if
			        sTemplate = fn_replace(sTemplate,"OBJ_ATTRIBUTE",sAttribute)
		        else
			        sTemplate = fn_replace(sTemplate,"OBJ_ATTRIBUTE","")
		        end if
		        sSpecialPrice=sSpecialPrice+aprice_dif
		        retail_price=retail_price+aprice_dif_retail
                  
       if instr(sTemplate,"OBJ_PRICE") or instr(sTemplate,"OBJ_SPECIAL_PRICE_OBJ") or instr(sTemplate,"OBJ_FINAL_PRICE_OBJ") or instr(sTemplate,"OBJ_MATRIX_OBJ") then
           
		        	if sSpecialPrice=retail_price then
				        sSpecialPrice=""
			        end if
			        
		        	sPrice = retail_price
		        	sFinalPrice = sPrice
		        	sFinalPrice1 = sPrice

		        	if sSpecialPrice<>"" then
			        sFinalPrice = sSpecialPrice
			        sSpecialPrice = sFinalPrice
		        	end if

		        	if sHide_Price then
			        sPrice = ""
			        sFinalPrice = ""
			        sSpecialPrice = ""
		        	elseif sEnable_Cust_Price then
		            if instr(sTemplate,"OBJ_FINAL_PRICE_OBJ")=0 then
			            sPrice = Store_Currency&"<input type=text name=cust_price value='"&sPrice&"' size=6>"
			        end if
    			    
			        sFinalPrice = Store_Currency&"<input type=text name=cust_price value='"&sFinalPrice&"' size=6>"
			        sSpecialPrice = ""
		        	else
		            sPrice = Currency_Format_Function(sPrice)
			        sFinalPrice = Currency_Format_Function(sFinalPrice)
			        sSpecialPrice = Currency_Format_Function(sSpecialPrice)
		        	end if
    		    
		        	if sSpecialPrice<>"" then
			        sPrice = "<strike>"&sPrice&"</strike>"
		        	end if
    		    
		        if instr(sTemplate,"OBJ_MATRIX_OBJ") then
		            if sUse_Price_By_Matrix and Service_Type => 5 then
		                sql_select="select Item_Price, Matrix_low, Matrix_high from Store_items_Price_Matrix WITH (NOLOCK) where Item_Id = "&Item_Id&" and store_id = "&store_id&" order by matrix_low"
			            set itemfields3=server.createobject("scripting.dictionary")
			            Call DataGetrows(conn_store,sql_select,itemdata3,itemfields3,noRecords3)
		        	    sMatrixText=""	
				        if noRecords3 = 0 then
					        FOR itemrowcounter3= 0 TO itemfields3("rowcount")
					            sMatrixHigh = itemdata3(itemfields3("matrix_high"),itemrowcounter3)
                                                    sMatrixLow = itemdata3(itemfields3("matrix_low"),itemrowcounter3)
                                                    sMatrixText = sMatrixText & "<BR>Buy "&sMatrixLow
						    sMatrixPrice = Currency_Format_Function(itemdata3(itemfields3("item_price"),itemrowcounter3)+aprice_dif)

                                                        If sMatrixHigh=-1 then
							        sMatrixText = sMatrixText & "&nbsp;or more items for "&sMatrixPrice&" each"
						        elseIf sMatrixHigh=sMatrixLow then
							        sMatrixText = sMatrixText & "&nbsp;items for "&sMatrixPrice&" each"
						        Else
							        sMatrixText = sMatrixText & "-"&sMatrixHigh&" items for "&sMatrixPrice&" each"
						        End If
					        Next
				        end if
			        Else
				        sMatrixText = ""
			        End If
			        sTemplate = fn_replace(sTemplate,"OBJ_MATRIX_OBJ",sMatrixText)
		        end if
		        end if

		        if instr(sTemplate,"OBJ_STOCK_OBJ") then
			        if Service_Type => 5 then
				        if sQuantity_Control  = -1 and sHide_Stock_This = 0 then
					        if stock_corection<>-1 then
						        if stock_corection<exStock then
							        exStock = stock_corection
						        Else
							        exStock = exStock
						        End If
					        Else
						        exStock = exStock
					        End If
					        if exStock>0 then
						        sStockText = "In Stock ("&exStock&")"
					        Else
						        sStockText = "Out Of Stock"
					        End If

				        Else

					        sStockText = ""
				        End If
			        end if
			        sTemplate = fn_replace(sTemplate,"OBJ_STOCK_OBJ",sStockText)
		        end if

		        If bSearch=1 or (sQuantity_Control = -1 and exStock<=sQuantity_Control_Number) then
			        sQuantityInput = ""
			        sOrderButton = ""
		        End If
    					
		        sTemplate = fn_replace(sTemplate,"OBJ_EXT_FIELD1_OBJ",M_d_1)
		        sTemplate = fn_replace(sTemplate,"OBJ_EXT_FIELD2_OBJ",M_d_2)
		        sTemplate = fn_replace(sTemplate,"OBJ_EXT_FIELD3_OBJ",M_d_3)
		        sTemplate = fn_replace(sTemplate,"OBJ_EXT_FIELD4_OBJ",M_d_4)
		        sTemplate = fn_replace(sTemplate,"OBJ_EXT_FIELD5_OBJ",M_d_5)

		        if instr(sTemplate,"OBJ_UD_OBJ") then
			        UD_Text = ""
			        if Service_Type => 5 and bSearch = 0 then
				        UD_Text = "<table>"
					        If use_u_d_1  <> 0 then
						        u_d_1 = fn_get_querystring("u_d_1_"&Item_Id)
						        User_Defined_Fields = User_Defined_Fields
						        if isNull(User_Defined_Fields) or User_Defined_Fields = "" then
							        UD_Text = UD_Text & "<tr><td>"&u_d_1_name&"<br><textarea name='U_d_1' cols='35' rows='2'>"&u_d_1&"</textarea><INPUT name=U_d_1_C type=hidden value='Op|String|0|400|||"&u_d_1_name&"'></td></tr>"
						        else
							        UD_Text = UD_Text & "<tr><td>"&u_d_1_name&"<br><input name='U_d_1' type=text "&User_Defined_Fields&" value='"&u_d_1&"'><INPUT name=U_d_1_C type=hidden value='Op|String|0|400|||"&u_d_1_name&"'></td></tr>"
						        end if
					        End If

					        If use_u_d_2 <> 0 then
						        User_Defined_Fields_2 = User_Defined_Fields_2
						        u_d_2 = fn_get_querystring("u_d_2_"&Item_Id)
						        if isNull(User_Defined_Fields_2) or User_Defined_Fields_2 = "" then
							        UD_Text = UD_Text & "<tr><td>"&u_d_2_name&"<br><textarea name='U_d_2' cols='35' rows='2'>"&u_d_2&"</textarea><INPUT name=U_d_2_C type=hidden value='Op|String|0|400|||"&u_d_2_name&"'></td></tr>"
						        else
							        UD_Text = UD_Text & "<tr><td>"&u_d_2_name&"<br><input name='U_d_2' type=text "&User_Defined_Fields_2&" value='"&u_d_2&"'><INPUT name=U_d_2_C type=hidden value='Op|String|0|400|||"&u_d_2_name&"'></td></tr>"
						        end if
					        End If

					        If use_u_d_3 <> 0 then
						        User_Defined_Fields_3 = User_Defined_Fields_3
						        u_d_3 = fn_get_querystring("u_d_3_"&Item_Id)
						        if isNull(User_Defined_Fields_3) or User_Defined_Fields_3 = "" then
							        UD_Text = UD_Text & "<tr><td>"&u_d_3_name&"<br><textarea name='U_d_3' cols='35' rows='2'>"&u_d_3&"</textarea><INPUT name=U_d_3_C type=hidden value='Op|String|0|400|||"&u_d_3_name&"'></td></tr>"
						        else
							        UD_Text = UD_Text & "<tr><td>"&u_d_3_name&"<br><input name='U_d_3' type=text "&User_Defined_Fields_3&" value='"&u_d_3&"'><INPUT name=U_d_3_C type=hidden value='Op|String|0|400|||"&u_d_3_name&"'></td></tr>"
						        end if
					        End If

					        If use_u_d_4 <> 0 then
						        User_Defined_Fields_4 = User_Defined_Fields_4
						        u_d_4 = fn_get_querystring("u_d_4_"&Item_Id)
						        if isNull(User_Defined_Fields_4) or User_Defined_Fields_4 = "" then
							        UD_Text = UD_Text & "<tr><td>"&u_d_4_name&"<br><textarea name='U_d_4' cols='35' rows='2'>"&u_d_4&"</textarea><INPUT name=U_d_4_C type=hidden value='Op|String|0|400|||"&u_d_4_name&"'></td></tr>"
						        else
							        UD_Text = UD_Text & "<tr><td>"&u_d_4_name&"<br><input name='U_d_4' type=text "&User_Defined_Fields_4&" value='"&u_d_4&"'><INPUT name=U_d_4_C type=hidden value='Op|String|0|400|||"&u_d_4_name&"'></td></tr>"
						        end if
					        End If
    						
					        If use_u_d_5 <> 0 then
						        User_Defined_Fields_5 = User_Defined_Fields_5
						        u_d_5 = fn_get_querystring("u_d_5_"&Item_Id)
						        if isNull(User_Defined_Fields_5) or User_Defined_Fields_5 = "" then
							        UD_Text = UD_Text & "<tr><td>"&u_d_5_name&"<br><textarea name='U_d_5' cols='35' rows='2'>"&u_d_5&"</textarea><INPUT name=U_d_5_C type=hidden value='Op|String|0|400|||"&u_d_5_name&"'></td></tr>"
						        else
							        UD_Text = UD_Text & "<tr><td>"&u_d_5_name&"<br><input name='U_d_5' type=text "&User_Defined_Fields_5&" value='"&u_d_5&"'><INPUT name=U_d_5_C type=hidden value='Op|String|0|400|||"&u_d_5_name&"'></td></tr>"
						        end if
					        End If

				        UD_Text = UD_Text & "</table>"
			        end if
			        sTemplate = fn_replace(sTemplate,"OBJ_UD_OBJ",UD_Text)
		        end if

		        sTemplate = fn_replace(sTemplate,"OBJ_DESCRIPTIONS_OBJ",sDescriptionS)
		        sTemplate = fn_replace(sTemplate,"OBJ_DESCRIPTIONL_OBJ",sDescriptionL)
		        sTemplate = fn_replace(sTemplate,"OBJ_DESCRIPTION_OBJ",sDescriptionL)



		        if instr(sTemplate,"OBJ_FRIEND") then
		        	sFriendURL = replace(sDetailUrl,"-detail.htm","-friend.htm")
			     sFriend = "<a href='"&sFriendURL&"' class=link>Send to a friend</a>"
			     sTemplate = fn_replace(sTemplate,"OBJ_FRIEND_OBJ",sFriend)
			     sTemplate = fn_replace(sTemplate,"OBJ_FRIEND_URL_OBJ",sFriendURL)
		        end if
                  
		        sTemplate = fn_replace(sTemplate,"OBJ_PRICE",sPrice)
		        sTemplate = fn_replace(sTemplate,"OBJ_SPECIAL_PRICE_OBJ",sSpecialPrice)

		        sTemplate = fn_replace(sTemplate,"OBJ_FINAL_PRICE_OBJ",sFinalPrice)

		        sTemplate = fn_replace(sTemplate,"OBJ_WEIGHT_OBJ",sItem_Weight)
		        sTemplate = fn_replace(sTemplate,"OBJ_HANDLING_FEE_OBJ",sItem_Handling)
		        sTemplate = fn_replace(sTemplate,"OBJ_RECURRING_FEE_OBJ",sRecurring_Fee)
		        sTemplate = fn_replace(sTemplate,"OBJ_RECURRING_DAYS_OBJ",sRecurring_Days)
		        sTemplate = fn_replace(sTemplate,"OBJ_SPECIAL_PRICE_DISCOUNT%_OBJ",sSpecial_Price_Discount)
		        sTemplate = fn_replace(sTemplate,"OBJ_SPECIAL_PRICE_START_DATE_OBJ",sSpecial_Start_Date)
		        sTemplate = fn_replace(sTemplate,"OBJ_SPECIAL_PRICE_END_DATE_OBJ",sSpecial_End_Date)
                  if inStr(sTemplate,"OBJ_QTY_OBJ") > 0 then
			        sTemplate = fn_replace(sTemplate,"OBJ_QTY_OBJ",sQuantityInput)
			     elseif instr(sTemplate,"quantity") = 0 and bSearch=0 then
				        sTemplate = sTemplate & sQuantityHidden
			     end if
		        	sTemplate = fn_replace(sTemplate,"OBJ_ORDER_OBJ",sOrderButton)
                  if instr(sTemplate,"OBJ_ACCESSORY_OBJ") and accessories_exist<>0 then
			        sAccessory = ""
			        if Service_Type => 5 then
				        sql_statement="wsp_item_accessory "&Store_Id&","&sHide_Out_stock&","&Item_id&",'"&Groups&"';"

				        item_f_layout = fn_replace(item_f_layout,"OBJ_IMAGE_OBJ","OBJ_IMAGES_OBJ")
                        item_f_layout = fn_replace(item_f_layout,"OBJ_DESCRIPTION_OBJ","OBJ_DESCRIPTIONS_OBJ")
                        sAccessory = fn_create_item(item_f_layout,sql_statement,1,sThisUrl,500,item_f_display,0)
				        sAccessory = Dict_Accessory&"<HR>"&sAccessory
			        end if
			        sTemplate = fn_replace(sTemplate,"OBJ_ACCESSORY_OBJ",sAccessory)
		        else
			        sTemplate = fn_replace(sTemplate,"OBJ_ACCESSORY_OBJ","")
		        end if
		        sWidth = cint(100/(item_display + 1))

		        if bSearch<>1 then
           
                  	sTemplate=sForm&sHiddenFields&sTemplate&"</form>"
		        end if
		        sItemText = sItemText & "<TD width='"&sWidth&"%'>"&sTemplate&"</td>"
		        'RELATE TO NAVIGATIONAL BAR
		        Relative_Row = Relative_Row + 1

		        itemrowcounter=itemrowcounter+1
		        j=j+1
		        i=oldi
    		
		        sFormName1 = "ITEM_"&Item_Id
		        sSubmitName1 = "Insert_To_Cart"

		        if bSearch =0 then

		            sItemText = sItemText & ("<SCRIPT language=""JavaScript"" type=""text/javascript"">"&vbcrlf&_
			            "var frmvalidator  = new Validator("""&sFormName1&""");"&vbcrlf&_
                        "frmvalidator.addValidation(""quantity"",""greaterthan="&sQuantityMin-.01&""",""Quantity must be at least "&sQuantityMin&""");"&vbcrlf&_
				        "frmvalidator.addValidation(""quantity"",""req"",""Please enter a quantity"");"&vbcrlf)
			        if sQuantity_Control =-1 then
				         exStockOrder = sQuantity_InStock - sQuantity_Control_Number
				         sItemText = sItemText & ("frmvalidator.addValidation(""quantity"",""lessthan="&exStockOrder+.01&""",""Either you have attempted to order more items than we have in stock or we have reached our low quantity and cannot guarantee shipment, please enter a quantity less than "&exStockOrder&"."");"&vbcrlf )
				   end if
			        if sFractional <> -1 then
					     sItemText = sItemText & ("frmvalidator.addValidation(""quantity"",""numeric"",""Quantity must be a number."");"&vbcrlf)
			        end if

				    sItemText = sItemText & ("</script>"&vbcrlf)
		        end if
			loop 
		
			sItemText = sItemText & ("</tr>")
			i=i+1
		loop
		sItemText = sItemText & ("</table>"&vbcrlf)

		set itemfields3 = Nothing
		set accfields = Nothing
	end if
	set itemfields = Nothing
	sItemText = sItemText & sNavigation & sPage_List

	fn_create_item = sItemText

end function

function fn_show_breadcrumbs (Show_TopNav,Full_Name,q_Item_Page_Name)
	sString=""
	if Show_TopNav then
	    sString = "<table border='0' cellspacing='0' width='100%'><tr><td>"
		if Top_Depts="" then
		    Top_Depts="Top"
		end if
		if Full_Name="" or Full_Name="Un Assigned" then
			sString = sString & ("<b>"&Top_Depts&"</b>")
		else
			sString = sString & ("<b><a href='"&fn_dept_url("","")&"' class=link>"&Top_Depts&"</a></b>")
		    sDeptString=""
		    for each sThisDeptName in split(Full_Name," > ")
			     if sDeptString = "" then
	 			     sDeptString = sThisDeptName
			     else
	 			      sDeptString = sDeptString & " > " & sThisDeptName
			     end if
			     if q_Item_Page_Name="" and sDeptString=Full_Name then
			        sString = sString & ("<b> > "&sThisDeptName&"</b>")
			     else
			        sString = sString & ("<b> > <a href='"&fn_dept_url(sDeptString,"")&"' class=link>"&sThisDeptName&"</a></b>")
			     end if
		    next
		end if

		sString = sString & ("<hr></td></tr></table>")
	else
		sString = ""
	end if
	fn_show_breadcrumbs=sString

end function

function fn_get_recordset_count (sql_count)
	fn_print_debug "sql_count="&sql_count
	set rs_store=conn_store.execute(sql_count)
	fn_get_recordset_count=rs_store(0)
	rs_store.close
end function

function fn_get_end_row (number_of_recordset,iCols,iRows,Start_Row,sType)
	if (iCols*iRows)>=number_of_recordset then
	    sNumberPerPage = (iCols*iRows-1) + 1
	elseif (Request.Form("Next_Jump_Click"&sType) <> "" or Request.Form("Next_Jump_Click"&sType&".x") <> "") and isNumeric(Request.Form("Next_jump"&sType)) then
	    sNumberPerPage = Request.Form("Next_jump"&sType)
	elseif (Request.Form("Back_Jump_Click"&sType) <> "" or Request.Form("Back_Jump_Click"&sType&".x") <> "") and isNumeric(Request.Form("Back_jump"&sType)) then
	    sNumberPerPage = Request.Form("Back_jump"&sType)
	else
	    sNumberPerPage = (iCols*iRows-1) + 1
	end if
	End_Row = sNumberPerPage + Start_Row

        if End_Row > number_of_recordset then
		End_Row = number_of_recordset
	end if
	fn_get_end_row = End_Row
end function

function fn_get_start_row (number_of_recordset,iCols,iRows,sType)
	if (iCols*iRows)>=number_of_recordset then
	    Start_Row = 0
	elseif fn_get_querystring("Jump_To_Page") <> "" and isNumeric(fn_get_querystring("Jump_To_Page")) then	
		Start_Row = fn_get_querystring("Jump_To_Page") * ((iCols*iRows))
		if Start_Row >= number_of_recordset then
			Start_Row = (number_of_recordset) - (iCols*iRows)
		end if
	elseif fn_get_querystring("repost")<>"" then
		Start = split(fn_get_querystring("start"),",")
		Start_Row = Start(0)
	Elseif (Request.Form("Next_Jump_Click"&sType) <> "" or Request.Form("Next_Jump_Click"&sType&".x") <> "") and isNumeric(Request.Form("Current_End_Row"&sType&sType)) then
		Start_Row = cint(Request.Form("Current_End_Row"&sType&sType))
	Elseif (Request.Form("Back_Jump_Click"&sType) <> "" or Request.Form("Back_Jump_Click"&sType&".x") <> "") and (isNumeric(Request.Form("Current_Start_Row"&sType&sType)) and isNumeric(Request.Form("Back_jump"&sType)) ) then
		Start_Row = cint(Request.Form("Current_Start_Row"&sType&sType)) -  Cint(Request.Form("Back_jump"&sType))
	Else
		Start_Row = 0
	End if
	
	if Start_Row < 0 or isNull(Start_Row) or Start_Row="" then
		Start_Row = 0
	End if
	fn_get_start_row=Start_Row
end function

function fn_paging (number_of_recordset,Start_Row,End_Row)
	if Start_Row>0 or End_Row < number_of_recordset then
		fn_paging=1
	else
		fn_paging=0
	end if
end function

function fn_make_navigation (sUrl,number_of_recordset,iCols,iRows,Start_Row,End_Row,sType)
	fn_print_debug "in make navigation "&sUrl

    if Request("Jump_To"&sType)<>"" then
        sJump_To = "Jump_To"&sType&"="&Request("Jump_To"&sType)
    end if
	sPostUrl = sUrl
	if sPostUrl = sUrl and sJump_To<>"" then
	    sPostUrl = sPostUrl&"?"&sJump_To
	elseif sJump_To<>"" then
	    sPostUrl = sPostUrl&"&"&sJump_To
	end if
	sRecords = sRecords & "<form action='"&sPostUrl&"' method='post'>"
	
	sRecords = sRecords & "<input type='hidden' name='Current_Start_Row"&sType&sType&"' value='"&Start_Row&"'>"
	sRecords = sRecords & "<input type='hidden' name='Current_End_Row"&sType&sType&"' value='"&End_Row&"'>"
	sRecords = sRecords & "<div align='center'><center>"
	sRecords = sRecords & "<table border='0' width='63%'><tr>"
	sNextButton = fn_create_action_button ("Button_image_Next", "Next_Jump_Click"&sType, "Next")
	sPrevButton = fn_create_action_button ("Button_image_Prev", "Back_Jump_Click"&sType, "Prev")
	If Start_Row>0 then
		sRecords = sRecords & "<td width='30%' align='center' class='normal'><b><input type='hidden' name='Back_jump"&sType&"' size='3' value='"&iCols*iRows&"'></b></td><td>"&sPrevButton&"</td>"
	Else
		sRecords = sRecords & "<td width='30%' align='center' class='normal'>&nbsp;</td>"
	End If
	sRecords = sRecords & "<td width='40%' align='center' nowrap class='normal'>Displaying "&Start_Row+1&" - "& End_Row&" of "& number_of_recordset &"&nbsp;</td>"
	If cint(End_Row) < cint(number_of_recordset) then
			sRecords = sRecords & "<td width='30%' align='center' class='normal'>"&sNextButton&"</td><td><b><input type='hidden' name='Next_jump"&sType&"' size='3' value='"&iCols*iRows&"'></b></td>"
	Else
		sRecords = sRecords & "<td width='30%' align='center' class='normal'>&nbsp;</td>"
	End If
	sRecords = sRecords & "</tr></table></center></div></form>"

	fn_make_navigation = sRecords&sJumpMenu

end function

function fn_make_navigation_item (sUrl,number_of_recordset,iCols,iRows,Start_Row,End_Row,sType)
	fn_print_debug "in make navigation "&sUrl&" "&sType

    if Request("Jump_To_Items")<>"" then
        sJump_To = "Jump_To"&sType&"="&Request("Jump_To_Items")
    end if
	sPostUrl = sUrl

	sRecordsPerPage = iCols * iRows
	sThisPage = round(Start_Row/sRecordsPerPage)

	sPageUrlNext = replace(sPagePattern,"%OBJ_PAGE_OBJ%",sThisPage+1)
	sPageUrlPrev = replace(sPagePattern,"%OBJ_PAGE_OBJ%",sThisPage-1)
	
	if sJump_To<>"" then
     	sPageUrlNext=fn_append_querystring(sPageUrlNext,sJump_To)
     	sPageUrlPrev=fn_append_querystring(sPageUrlPrev,sJump_To)
     	sPostUrl=fn_append_querystring(sPostUrl,sJump_To)
	end if

	sRecords = sRecords & "<div align='center'><center>"
	sRecords = sRecords & "<table border='0' width='63%'><tr>"
	sNextButton = fn_create_action_button ("Button_image_Next", "Next", "Next")
	sPrevButton = fn_create_action_button ("Button_image_Prev", "Prev", "Prev")
	If Start_Row>0 then
		sRecords = sRecords & "<form action='"&sPageUrlPrev&"' method='post'>"
		sRecords = sRecords & "<td width='30%' align='center' class='normal'></td><td>"&sPrevButton&"</td></form>"
	Else
		sRecords = sRecords & "<td width='30%' align='center' class='normal'>&nbsp;</td>"
	End If
	sRecords = sRecords & "<td width='40%' align='center' nowrap class='normal'>Displaying "&Start_Row+1&" - "& End_Row&" of "& number_of_recordset &"&nbsp;</td>"
	If cint(End_Row) < cint(number_of_recordset) then
		sRecords = sRecords & "<form action='"&sPageUrlNext&"' method='post'>"
		sRecords = sRecords & "<td width='30%' align='center' class='normal'>"&sNextButton&"</td><td></td></form>"
	Else
		sRecords = sRecords & "<td width='30%' align='center' class='normal'>&nbsp;</td>"
	End If
	sRecords = sRecords & "</tr></table></center></div>"

	fn_make_navigation_item = sRecords&sJumpMenu

end function

function fn_jump_menu (sLink)
    fn_print_debug "in jump menu"
	if Show_Jump then
		sJumpItems = "<table border='0' width='100%' cellspacing='0' align='center'><tr>"&_
			"<td width='100%' class='normal' colspan=6 align=center>Jump To:"
		sAlphabet = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
		for each sLetter in split(sAlphabet,",")
			sJumpItems = sJumpItems & ("<a class='link' href='"&sLink&"?Jump_To_Items="&sLetter&"'>"&sLetter&"</a> ")
		next
		sJumpItems = sJumpItems & ("<a class='link' href='"&sLink&"'>All</a></td></tr></table>")
		fn_jump_menu = sJumpItems
	end if
end function

function fn_jump_menu1 (sLink,sSelectedLetter)
    fn_print_debug "in jump menu2 "&sSelectedLetter
	if Show_Jump then
		sJumpItems = "<table border='0' width='100%' cellspacing='0' align='center'><tr>"&_
			"<td width='100%' class='normal' colspan=6 align=center>Jump To:"
		sAlphabet = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
		for each sLetter in split(sAlphabet,",")
			if sSelectedLetter=sLetter then
				sJumpItems = sJumpItems & ("<B>"&sLetter&"</B> ")
			else
				sJumpItems = sJumpItems & ("<A class='link' href='"&sLink&"?Jump_To_Items="&sLetter&"'>"&sLetter&"</A> ")
			end if
		next
		sJumpItems = sJumpItems & ("<a class='link' href='"&sLink&"'>All</a></td></tr></table>")
		fn_jump_menu1 = sJumpItems
	end if
end function

function fn_make_pagelist (sFirstPage,sPagePattern,number_of_recordset,iCols,iRows)
	fn_print_debug "in make pagelist icols="&icols&",irows="&irows
	
	sPages = cint(number_of_recordset/(iCols*iRows)+.49)
	sPageList = "<table width='100%'><tr><td align=center width='100%'>Page:"
	FOR iPage= 0 TO sPages - 1
	    if iPage=0  then
	        sPageUrl = sFirstPage
	    else
	        sPageUrl = replace(sPagePattern,"%OBJ_PAGE_OBJ%",iPage)
	    end if
		    if iPage>0 then
		    
		    end if
			sLink = "<a class=link href='"&sPageUrl&"'>"
		sPageList = sPageList & " " & sLink & (iPage + 1) & "</a>"
	next
	sPageList = sPageList & "</td></tr></table>"
	fn_make_pagelist=sPageList
end function 

function fn_get_item_count(sDepartment_Id)
	Set rs_store5 = Server.CreateObject("ADODB.Recordset")
	sql_count="exec wsp_items_count_browse "&store_Id&","&sHide_Out_stock&","&sDepartment_Id&",'';"
    fn_print_debug sql_count
	rs_Store5.open sql_count,conn_store, 1, 1
	if not rs_Store5.eof then
		if not isnull(rs_store5(0)) then
			fn_get_item_count = clng(rs_store5(0))
		else
			fn_get_item_count = 0
		end if
	else
		fn_get_item_count = 0
	End if
	rs_Store5.close
	Set rs_store5 = nothing
end function

%>
<!--#include file="items_attr_sub.asp"-->
