<% 'ATTRIBUTES BUILDING

sAttribute_class=""
sAttributeId_Sel = fn_get_querystring("Attributes")
if not(isNumeric(sAttributeId_Sel)) then
	sAttributeId_Sel=""
end if
for each key in request.querystring
    key=fn_sanitize(key)
    if instr(key,"Count_Class_")=1 then

        sValueArray=split(fn_get_querystring(key),"#")
	   if isNumeric(sValueArray(0)) then
	   		if sAttributeId_Sel<>"" then
            		sAttributeId_Sel=sAttributeId_Sel&","
        		end if
        		sAttributeId_Sel=sAttributeId_Sel&sValueArray(0)
     	end if
    end if
next

'sql_attributes = "SELECT attribute_class,attribute_type,attribute_id,use_item,iitem_id,iitem_quantity,attribute_price_difference,attribute_weight_difference, attribute_value,attribute_sku,attribute_hidden,attribute_value    FROM store_items_attributes WITH (NOLOCK) left join dbo.SplitIDs('"&sAttributeId_Sel&"') as sSplit on attribute_id=sSplit.sId WHERE Store_id="&Store_id&" AND Item_Id="&Item_Id&" ORDER BY Attribute_class,sId desc, [Default] desc,Display_Order, Attribute_value;"
sql_attributes = "exec wsp_item_attribute "&store_id&","&item_id&",'"&sAttributeId_Sel&"';"
fn_print_debug sql_attributes
set rs2fields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_attributes,rs2data,rs2fields,noRecords2)

'CHECK IF THERE ARE ATTRIBUTES 
Count_Class =0

sAttribute = "<table border='0'>"
'LOOP OVER THE ATTRIBUTES CLASS 
Price_Add_Attr = 0
Set rs_storelim = Server.CreateObject("ADODB.Recordset") 

cattr = rs2fields("rowcount")

													
redim attrAll(cattr)
for i=0 to cattr
	attrAll(i)=", " 
next

sAttribute_class_old=""
sAttribute_class=""
	
on error resume next
if sConfigs_Exist then
	sql_configs = "select * from Store_Items_Configs WITH (NOLOCK) where store_id="&Store_id&" and item_id="&Item_Id
	fn_print_debug sql_configs
	set cfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_configs,cdata,cfields,noRecordsc)

	if noRecordsc = 0 then
		FOR rowcounterc= 0 TO cfields("rowcount")
			sel_First = -1
			sel_Sel = -1
			If cdata(cfields("config_desc"),rowcounterc)<>"" then
				sAttribute = sAttribute & "<tr><td colspan='3' class='normal'>"&cdata(cfields("config_desc"),rowcounterc)&"</td></tr>"
			End If

			sAttribute = sAttribute & "<tr><td nowrap class='normal'>"&cdata(cfields("config_name"),rowcounterc)&"</td>"&_
				"<td nowrap class='normal'>&nbsp;</td>"
			sql_opt = "select distinct Name from Store_Items_Configs_Dets WITH (NOLOCK) where store_id="&store_id&" and item_id="&Item_Id&" and config_id="&cdata(cfields("config_id"),rowcounterc)
			fn_print_debug sql_opt
			set optfields=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_opt,optdata,optfields,optRecords)

			selectName = "CNF_"&Item_Id&"_"&cdata(cfields("config_id"),rowcounterc)
			theUrl=""
			for each key in request.queryString
				key=fn_sanitize(key)
				if key<>"dept_page_name" and key<>"item_page_name" and key<>"jump_to_page" and key<>selectName and inStr(key, "Count_Class")<=0 and inStr(key, "Attributes")<=0 and instr(key,"start")<=0 and instr(key,"end")<=0 and instr(key,"repost")<=0 and instr(key,"u_d_1")<=0  and instr(key,"u_d_2")<=0 and instr(key,"u_d_3")<=0  and instr(key,"u_d_4")<=0  and instr(key,"u_d_5")<=0 then
					if theUrl<>"" then
						theUrl = theUrl&"&"&key&"="&fn_get_querystring(key)
					else 
						theUrl = key&"="&fn_get_querystring(key)
					end if
				end if 
			next
				
			sAttribute = sAttribute & "<td class='normal'>"&_
			"<select name='CNF_"&Item_Id&"_"&cdata(cfields("config_id"),rowcounterc)&"' onChange=""JavaScript:goReload('ITEM_"&Item_Id&"','dropdown', '"&selectName&"', '"&Item_Id&"', '"&theUrl&"', '"&theFileName&"');"">"
			if noRecordsopt = 0 then
				FOR rowcounteropt= 0 TO optfields("rowcount")
					optID = getConfOptID(Item_Id, cdata(cfields("config_id"),rowcounterc), optdata(optfields("name"),rowcounteropt))
					if sel_First = -1 then 
						sel_First = optID 
					end if
					sAttribute = sAttribute & "<option "
						if fn_get_querystring("CNF_"&Item_Id&"_"&cdata(cfields("config_id"),rowcounterc))=cstr(optID) then
							sel_Sel = optID
							sAttribute = sAttribute & "selected "
						end if
							sAttribute = sAttribute & "value='"&optID&"'>"&optdata(optfields("name"),rowcounteropt)
					sAttribute = sAttribute & "</option>"
					if sel_sel = -1 then
						sel_sel = sel_First
					end if
				Next
			End If
			sAttribute = sAttribute & "</select></td></tr>"
	
			sel_sel_name = getConfOptName(Item_Id, cdata(cfields("config_id"),rowcounterc), sel_sel)
			sql_sel_attribs = "select distinct attribute_class from store_items_attributes WITH (NOLOCK) where store_id="&Store_id&" and item_id="&Item_Id
			fn_print_debug sql_sel_attribs
			set attfields=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_sel_attribs,attdata,attfields,noRecordsatt)

			cattr = 0 
			if noRecordsatt = 0 then
				FOR rowcounteratt= 0 TO attfields("rowcount")
                    sAttribute_ClassC=attdata(attfields("attribute_class"),rowcounteratt)
					sql_sel_lim = "select Attributes_Values from Store_Items_Configs_Dets WITH (NOLOCK) where store_id="&store_id&" and item_id="&Item_Id&" and config_id="&cdata(cfields("config_id"),rowcounterc)&" and name='"&sel_sel_name&"' and Attributes_Class='"&sAttribute_ClassC&"'"
					fn_print_debug sql_sel_lim
					rs_storelim.open sql_sel_lim, conn_store, 1, 1 
					if not rs_storelim.eof then 
						lims = rs_storelim("Attributes_Values") 
						attrAll(cattr) = attrAll(cattr)+lims+", " 
						sIndex="attribute_class:"&sAttribute_ClassC
						rs2fields(sIndex)=rs2fields(sIndex)&lims&", " 
						fn_print_debug "rs2fields("&sIndex&")="&rs2fields(sIndex)
					End If 
					rs_storelim.close
					cattr = cattr + 1 
				Next
			End if
			sAttribute = sAttribute & "<tr><td colspan='3' class='normal'>&nbsp;</td></tr>"
		Next
	End If

	set cfields = Nothing
	Set rs_storelim = nothing
	Set optfields = Nothing
	Set attfields = Nothing	
end if

if fn_get_querystring("Attributes") <> "" then
	attribute_ids = split(fn_get_querystring("Attributes"),",")
	num_attribute_ids = ubound(attribute_ids)
end if

'BUILD THE DROP DOWN MENU WITH THE ATTRIBUTES 


if noRecords2 = 0 then
	Count_Class=0
	FOR rowcounter2= 0 TO rs2fields("rowcount")
	    sAttribute_id=trim(rs2data(rs2fields("attribute_id"),rowcounter2))
	    sAttribute_class=trim(rs2data(rs2fields("attribute_class"),rowcounter2))
		sConfigMatch=rs2fields("attribute_class:"&sAttribute_class)
		if not(sConfigs_Exist) or sConfigMatch="" or Is_In_Collection(sConfigMatch,sAttribute_id,",") then
		    if sAttribute_class<>sAttribute_class_old then
		        sAttribute_class_old=sAttribute_class
		        sType=rs2data(rs2fields("attribute_type"),rowcounter2)
		        sFirst=1
		        attribute_Price_Difference_sel=0
			    attribute_Weight_Difference_sel=0
			    attribute_Retail_Difference_sel=0
			    Count_Class = Count_Class + 1
		        theUrl=""
		        selectName = "Count_Class_"&Count_Class&"_"&Item_Id
		        for each key in request.queryString
		        	   key=fn_sanitize(key)
			        if key<>"Attributes" and key<>"dept_page_name" and key<>"item_page_name" and key<>"jump_to_page" and key<>selectName and instr(key,"start")<=0 and instr(key,"end")<=0 and instr(key,"repost")<=0 and instr(key,"u_d_1")<=0  and instr(key,"u_d_2")<=0 and instr(key,"u_d_3")<=0	and instr(key,"u_d_4")<=0 and instr(key,"u_d_5")<=0 then
				        if theUrl<>"" then
					        theUrl = theUrl&"&"&key&"="&fn_get_querystring(key)
				        else 
					        theUrl = key&"="&fn_get_querystring(key)
				        end if
			        end if
		        next 
		        if sType="dropdown" then
				    sAttribute = sAttribute & "<tr><td nowrap class='normal'>Select "&sAttribute_class&"</td>"&_
					    "<td class='normal'>&nbsp;</td><td class='normal'>"&_
					    "<SELECT name='Count_Class_"&Count_Class&"' size='1' "
                    fn_print_debug "Reload_Attr="&Reload_Attr
			     if Reload_Attr then
                        sAttribute=sAttribute&"%RELOAD%"
                        sReload="onChange=""JavaScript:goReloadAttr('ITEM_"&Item_Id&"','dropdown', 'Count_Class_"&Count_Class&"', '"&Item_Id&"', '"&theUrl&"', '"&theFileName&"');"""
			        end if
                    sAttribute = sAttribute & ">"
                    sSelectedText="selected"
			        sBeforeSelected = "<OPTION "
			        sAfterSelected = "</OPTION>"
			        sEndRow = "</select></td></tr>"
		        else
			        sSelectedText="checked"
			        sAttribute=sAttribute&"<TR><TD>"&sAttribute_class&"</td></tr>"
			        sBeforeSelected = "<tr><td><input class='image' type='radio' name='Count_Class_"&Count_Class&"' "
                    if Reload_Attr then
                        sBeforeSelected=sBeforeSelected&"%RELOAD%"
                        sReload="onClick=""JavaScript:goReloadAttr('ITEM_"&Item_Id&"','radio', 'Count_Class_"&Count_Class&"', '"&Item_Id&"', '"&theUrl&"', '"&theFileName&"');"""
			        end if
                    sAfterSelected = "</td></tr>"
			        sEndRow = "</td></tr>"		
		        end if
		    else
		        sFirst=0
		    end if
		
		    use_item=rs2data(rs2fields("use_item"),rowcounter2)	
            
            if use_item<>0 then
                IItem_Id = rs2data(rs2fields("iitem_id"),rowcounter2)
                IItem_Quantity = rs2data(rs2fields("iitem_quantity"),rowcounter2)		
                sql_item = "select item_weight, retail_price, dbo.wf_get_price("&store_id&","&IItem_Id&",getdate(),-1,i.use_price_by_matrix,i.retail_price,i.Retail_Price_special_Discount,i.Special_start_date,i.Special_end_date,'"&Groups&"','',-1,'') as item_price FROM store_items i WITH (NOLOCK) WHERE i.store_id="&store_id&" AND i.item_id="&IItem_Id
                fn_print_debug sql_item
                rs_store.open sql_item,conn_store,1,1
                if not rs_store.eof then
	                attribute_Price_Difference = rs_store("Item_Price")*IItem_Quantity
	                attribute_Retail_Difference = rs_store("retail_price")*IItem_Quantity
	                attribute_Weight_Difference = rs_store("Attribute_Weight")*IItem_Quantity
                end if
                rs_store.close
            else
                attribute_Price_Difference = rs2data(rs2fields("attribute_price_difference"),rowcounter2)
                attribute_Weight_Difference = rs2data(rs2fields("attribute_weight_difference"),rowcounter2)
                attribute_Retail_Difference = attribute_Price_Difference
            end if   
						
		    if not isNumeric(attribute_Price_Difference) or attribute_Price_Difference=Null then
		       attribute_Price_Difference=0
		       attribute_Retail_Difference=0
		    end if 
            if sFirst=1 then
                attribute_Price_Difference_sel = attribute_Price_Difference
                attribute_Weight_Difference_sel = attribute_Weight_Difference
                attribute_Retail_Difference_sel = attribute_Retail_Difference
            end if
        
	        sPriceDiff=attribute_Price_Difference-attribute_Price_Difference_sel
	        sRetailDiff=attribute_Retail_Difference-attribute_Retail_Difference_sel

	        sAttribute = sAttribute & sBeforeSelected
	        if sFirst=1 then
		        sAttribute = sAttribute & sSelectedText
	        else
	            sAttribute=replace(sAttribute,"%RELOAD%",sReload)
	        End If
	        sAttribute = sAttribute & " value='"&sAttribute_Id&"|"&sAttribute_class&"="&rs2data(rs2fields("attribute_value"),rowcounter2)&"|"&attribute_Price_Difference&"|"&attribute_Weight_Difference&"|"&rs2data(rs2fields("attribute_sku"),rowcounter2)&"|"&rs2data(rs2fields("attribute_hidden"),rowcounter2)&"'>"&rs2data(rs2fields("attribute_value"),rowcounter2)&"&nbsp;"
	        if sPriceDiff > 0 then
		        sAttribute = sAttribute & "(+"&Currency_Format_Function(sPriceDiff)&")"
	        ElseIf sPriceDiff < 0 then 
		        sAttribute = sAttribute & "(-"&Currency_Format_Function(-sPriceDiff)&")"
	        End if
	        sAttribute = sAttribute & sAfterSelected

	        if sFirst=1 then
		        aprice_dif=aprice_dif+attribute_Price_Difference_sel
		        aprice_dif_retail=aprice_dif_retail+attribute_Retail_Difference_sel
	        end if
	    end if
    Next
End if

if (rs2fields("rowcount")=rowcounter2+1) or (sAttribute_class<>trim(rs2data(rs2fields("attribute_class"),rowcounter2+1))) then
	sAttribute = sAttribute & sEndRow
end if


set rs2fields = Nothing

sAttribute = sAttribute & "</table>"
sAttribute = sAttribute & "<input type='hidden' name='Count_Class' value='"&Count_Class&"' >"
%>
