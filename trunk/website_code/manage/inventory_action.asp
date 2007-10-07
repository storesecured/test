<!--#include file="Global_Settings.asp"-->


<% 

if not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if


'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%> <!--#include virtual="common/Error_Template.asp"--><% 
else
	select case Request.Form("Form_Name")

	' ADD A NEW ITEM
		'EDIT AN ITEM
		case "Item_Edit","Item_Basic_Edit"
			'RETRIVING FORM DATA
			if Request.Form("Show_homepage") <> "" then
				Show_homepage = Request.Form("Show_homepage")
			else
				Show_homepage = 0
			end if
			if Request.Form("Hide_Price") <> "" then
				Hide_Price = Request.Form("Hide_Price")
			else
				Hide_Price = 0
			end if
			if Request.Form("Cust_Price") <> "" then
				Cust_Price = Request.Form("Cust_Price")
			else
				Cust_Price = 0
			end if
			if Request.Form("Hide_Stock") <> "" then
				Hide_Stock = Request.Form("Hide_Stock")
			else
				Hide_Stock = 0
			end if
			if Request.Form("Use_Price_By_Matrix") <> "" then 
				Use_Price_By_Matrix = Request.Form("Use_Price_By_Matrix")
			else
				Use_Price_By_Matrix = 0
			end if
			if Request.Form("Taxable") <> "" then
				Taxable = Request.Form("Taxable")
			else 
				Taxable = 0
			end if
			if Request.Form("Fractional") <> "" then
				Fractional = Request.Form("Fractional")
			else 
				Fractional = 0
			end if

			if Request.Form("Item_pin") <> "" then
				Item_pin = Request.Form("Item_pin")
			else 
				Item_pin = 0
			end if

			Department_Id = Request.Form("Department_Id")
			Old_Department_Id = Request.Form("Old_Department_Id")
			Item_Id = Request.Form("Item_Id")
			Retail_Price = Request.Form("Retail_Price")
			Wholesale_Price = Request.Form("Wholesale_Price")
			Item_Weight = Request.Form("Item_Weight")
			Item_Handling = Request.Form("Item_Handling")
			Recurring_days = Request.Form("Recurring_days")
			Recurring_fee = Request.Form("Recurring_fee")
			View_Order = Request.Form("View_Order")
			Ship_Location_Id = Request.form("Ship_Location_Id")

			Quantity_in_stock = Request.Form("Quantity_in_stock")
			if Request.Form("Quantity_Control") <> "" then 
				Quantity_Control = Request.Form("Quantity_Control")
			else 
				Quantity_Control = 0
			end if
			Waive_Shipping = Request.Form("Waive_Shipping")
			if Request.Form("Waive_Shipping") <> "" then
				Waive_Shipping = Request.Form("Waive_Shipping")
			else 
				Waive_Shipping = 0
			end if
			Quantity_Control_Number = Request.Form("Quantity_Control_Number")
			Quantity_Minimum = Request.Form("Quantity_Minimum")
			Retail_Price_special_Discount = Request.Form("Retail_Price_special_Discount")
			Shipping_Fee = Request.Form("Shipping_Fee")

			if Request.Form("Show") <> "" then 
				Show = Request.Form("Show")
			else 
				Show = 0
			end if

			if isNumeric(View_Order) and isNumeric(Recurring_days) and isNumeric(Recurring_fee) and isNumeric(Show_homepage) and isNumeric(Use_Price_By_Matrix) and isNumeric(Taxable) and isNumeric(Fractional) and isNumeric(Item_Id) and isNumeric(Retail_Price) and isNumeric(Wholesale_Price) and isNumeric(Item_Weight) and isNumeric(Quantity_in_stock) and isNumeric(Quantity_Control) and isNumeric(Quantity_Control_Number) and isNumeric(Quantity_Minimum) and isNumeric(Retail_Price_special_Discount) and isNumeric(Shipping_Fee) and isNumeric(Show) then
				Item_Name = checkStringForQ(Request.Form("Item_Name"))
				Item_Page_Name = checkStringForQ(Request.Form("Item_Page_Name"))
				Item_Remarks = checkStringForQ(Request.Form("Item_Remarks"))
				Item_Sku = checkStringForQ(Request.Form("Item_Sku"))
				ImageS_Path = checkStringForQ(Request.Form("ImageS_Path"))
				ImageL_Path = checkStringForQ(Request.Form("ImageL_Path"))
				Description_S = nullifyQ(Request.Form("Description_S"))
				Description_L = nullifyQ(Request.Form("Description_L"))

 			    Custom_Link = checkStringForQ(Request.Form("Custom_Link"))

				Meta_Keywords = checkStringForQ(Request.Form("Meta_Keywords"))
				Meta_Description = checkStringForQ(Request.Form("Meta_Description"))
                	Meta_Title = checkStringForQ(Request.Form("Meta_Title"))
                	
                	Brand = checkStringForQ(Request.Form("Brand"))
                	Condition = checkStringForQ(Request.Form("Condition"))
                	Product_Type = checkStringForQ(Request.Form("Product_Type"))

				Special_start_date = checkStringForQ(Request.Form("Special_start_date"))
				Special_end_date = checkStringForQ(Request.Form("Special_end_date"))
	
				File_Location = checkStringForQ(Request.Form("File_Location"))

				U_d_1_name = checkStringForQ(Request.Form("U_d_1_name"))
				if Request.Form("U_d_1") <> "" then 
					U_d_1 = Request.Form("U_d_1")
				else
					U_d_1 = 0
				end if
	
				U_d_2_name = checkStringForQ(Request.Form("U_d_2_name"))
				if Request.Form("U_d_2") <> "" then 
					U_d_2 = Request.Form("U_d_2")
				else 
					U_d_2 = 0
				end if
	
				U_d_3_name = checkStringForQ(Request.Form("U_d_3_name"))
				if Request.Form("U_d_3") <> "" then 
					U_d_3 = Request.Form("U_d_3")
				else 
					U_d_3 = 0
				end if
	
				U_d_4_name = checkStringForQ(Request.Form("U_d_4_name"))
				if Request.Form("U_d_4") <> "" then 
					U_d_4 = Request.Form("U_d_4")
				else 
					U_d_4 = 0
				end if
				
				U_d_5_name = checkStringForQ(Request.Form("U_d_5_name"))
				if Request.Form("U_d_5") <> "" then
					U_d_5 = Request.Form("U_d_5")
				else 
					U_d_5 = 0
				end if
				


			'---------------------------------------Extended Fields -------------------------
				if Request.Form("M_d_1") <> "" then
					M_d_1 = nullifyQ(Request.Form("M_d_1"))
				else
					M_d_1 = ""
				end if
	
				if Request.Form("M_d_2") <> "" then 
					M_d_2 = nullifyQ(Request.Form("M_d_2"))
				else
					M_d_2 = ""
				end if

				if Request.Form("M_d_3") <> "" then 
					M_d_3 = nullifyQ(Request.Form("M_d_3"))
				else 
					M_d_3 = ""
				end if
	
				if Request.Form("M_d_4") <> "" then 
					M_d_4 = nullifyQ(Request.Form("M_d_4"))
				else 
					M_d_4 = ""
				end if
				
				if Request.Form("M_d_5") <> "" then
					M_d_5 = nullifyQ(Request.Form("M_d_5"))
				else 
					M_d_5 = ""
				end if
			'---------------------------------------Extended Fields -------------------------

	
				'UPDATE INVENTORY
            if view_order="" then
               view_order=0
            end if
            if wholesale_price="" then
               wholesale_price=0
            end if
            
            if Request.Form("Form_Name")="Item_Basic_Edit" then
                sql_update = "sp_item_basic_update "&store_id&","&Item_Id&",'"&Old_Department_Id&"','"&Department_Id&"','"&Item_Name&"','"&Item_Page_Name&"','"&Item_Sku&"',"&Retail_Price&",'"&Images_Path&"','"&Imagel_Path&"','"&Description_L&"'"
                
			 conn_store.execute(sql_update)


				Response.Redirect "edit_items.asp"
		    else
                if Item_id=0 then

				    sql_update = "insert into store_items (Store_Id,Brand,Condition,Product_type,Waive_Shipping,Ship_Location_Id,View_Order,Recurring_Days,Recurring_Fee, Meta_Keywords, Meta_Description, Meta_Title, Hide_Stock, Show_Homepage, Hide_Price, Enable_Cust_Price, Sys_Modified,Item_pin,	Fractional,	Item_Name,	Item_Page_Name, Item_Sku, Retail_Price, Wholesale_Price, ImageS_Path, 	ImageL_Path, Taxable, Use_Price_By_Matrix, U_d_1, U_d_1_name, U_d_2, U_d_2_name, U_d_3, U_d_3_name, U_d_4, U_d_4_name, U_d_5, U_d_5_name, Description_S,	  Description_L,   Item_Weight,   Item_Handling, Quantity_in_stock,   Quantity_Control,  Quantity_Control_Number, Quantity_Minimum, Retail_Price_special_Discount,   Special_start_date, Special_end_date,   Shipping_Fee, File_Location,   Show, Item_Remarks, M_d_1, M_d_2,  M_d_3, M_d_4,  M_d_5, Custom_Link) values ("&Store_Id&",'"&Brand&"','"&Condition&"','"&Product_Type&"',"&Waive_Shipping&","&Ship_Location_Id&","&View_Order&", "&Recurring_days&", "&Recurring_Fee&", '"&Meta_Keywords&"', '"&Meta_Description&"', '"&Meta_Title&"', "&Hide_Stock&", "&Show_Homepage&", "&Hide_Price&", "&Cust_Price&", '"&now()&"',"&Item_pin&",	"&Fractional&",	'"&Item_Name&"','"&Item_Page_Name&"',	'"&Item_Sku&"', "&Retail_Price&", "&Wholesale_Price&", '"&ImageS_Path&"', 	'"&ImageL_Path&"', "&Taxable&", "&Use_Price_By_Matrix&", "&U_d_1&", '"&U_d_1_name&"', "&U_d_2&", '"&U_d_2_name&"', "&U_d_3&", '"&U_d_3_name&"', "&U_d_4&", '"&U_d_4_name&"', "&U_d_5&", '"&U_d_5_name&"', '"&Description_S&"',	  '"&Description_L&"',   "&Item_Weight&",   "&Item_Handling&", "&Quantity_in_stock&",   "&Quantity_Control&",  "&Quantity_Control_Number&", "&Quantity_Minimum&", "&Retail_Price_special_Discount&",   '"&GetUSFormatDate(Special_start_date)&"', '"&GetUSFormatDate(Special_end_date)&"',   "&Shipping_Fee&", '"&File_Location&"',   "&Show&", '"&Item_Remarks&"', '"&M_d_1&"', '"&M_d_2&"',  '"&M_d_3&"', '"&M_d_4&"',  '"&M_d_5&"', '"&Custom_Link&"')"
				    conn_store.execute(sql_update)
					sql_select = "select max(item_id) from store_items where store_id="&store_id
					rs_store.open sql_select,conn_store,1,1
						Item_id = rs_store(0)
					rs_store.close

                else
			        sql_update = "update Store_Items set Brand='"&Brand&"',Condition='"&Condition&"',Product_Type='"&Product_Type&"',Waive_Shipping="&Waive_Shipping&",Ship_Location_Id="&Ship_Location_Id&",View_Order="&View_Order&", Recurring_Days="&Recurring_days&", Recurring_Fee="&Recurring_Fee&", Meta_Keywords='"&Meta_Keywords&"', Meta_Description='"&Meta_Description&"', Meta_Title='"&Meta_Title&"', Hide_Stock="&Hide_Stock&", Show_Homepage="&Show_Homepage&", Hide_Price="&Hide_Price&", Enable_Cust_Price="&Cust_Price&", Item_pin="&Item_pin&",	Fractional="&Fractional&",	Item_Name = '"&Item_Name&"',	Item_Page_Name='"&Item_Page_Name&"', Item_Sku = '"&Item_Sku&"', Retail_Price = "&Retail_Price&", Wholesale_Price = "&Wholesale_Price&", ImageS_Path = '"&ImageS_Path&"', 	ImageL_Path = '"&ImageL_Path&"', Taxable = "&Taxable&", Use_Price_By_Matrix = "&Use_Price_By_Matrix&", U_d_1 = "&U_d_1&", U_d_1_name = '"&U_d_1_name&"', U_d_2 = "&U_d_2&", U_d_2_name = '"&U_d_2_name&"', U_d_3 = "&U_d_3&", U_d_3_name = '"&U_d_3_name&"', U_d_4 = "&U_d_4&", U_d_4_name = '"&U_d_4_name&"', U_d_5 = "&U_d_5&", U_d_5_name = '"&U_d_5_name&"', Description_S = '"&Description_S&"',	  Description_L = '"&Description_L&"',   Item_Weight = "&Item_Weight&",   Item_Handling="&Item_Handling&", Quantity_in_stock = "&Quantity_in_stock&",   Quantity_Control = "&Quantity_Control&",  Quantity_Control_Number = "&Quantity_Control_Number&", Quantity_Minimum = "&Quantity_Minimum&", Retail_Price_special_Discount = "&Retail_Price_special_Discount&",   Special_start_date = '"&GetUSFormatDate(Special_start_date)&"', Special_end_date = '"&GetUSFormatDate(Special_end_date)&"',   Shipping_Fee = "&Shipping_Fee&", File_Location = '"&File_Location&"',   Show = "&Show&", Item_Remarks = '"&Item_Remarks&"', M_d_1 = '"&M_d_1&"', M_d_2 = '"&M_d_2&"',  M_d_3 = '"&M_d_3&"', M_d_4 ='"&M_d_4&"',  M_d_5 = '"&M_d_5&"', Custom_Link = '"&Custom_Link&"' WHERE Store_id="&Store_id&" AND Item_Id = "&Item_Id
	                	session("sql") = sql_update
	           
				    conn_store.Execute sql_update
		        end if
		        
		        if Old_Department_Id<>Department_Id then
                    sql_update2 = "sp_item_dept_update "&store_id&","&Item_Id&",'"&Department_Id&"'"
                    conn_store.execute(sql_update2)
                end if
                Response.Redirect "Item_edit.asp?Item_Id="&Item_Id&"&"&sAddString
	        end if
				Response.Redirect "Item_edit.asp?Item_Id="&Item_Id&"&"&sAddString
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if


		'ADD / EDIT AN ITEM ATTRIBUTE

		'========== RADIO BUTTONS IN CURRENT ATTRIBUTES ===========
		Case "Add_Attribute"
			Item_Id = Request.Form("Item_Id")

			Attribute_Type = Request.Form("Attribute_Type")

			Attribute_Price_difference = Request.Form("Attribute_Price_difference")
			Attribute_Weight_difference = Request.Form("Attribute_Weight_difference")

			Use_Item = Request.Form("Use_Item")
			IItem_ID = Request.Form("IItem_ID")

			IItem_Quantity = Request.Form("IItem_Quantity")

			if Request.Form("Default") <> "" then 
				Default = -1
			else
				Default = 0
			end if
			Display_Order = Request.Form("Display_Order")

			if isNumeric(Display_Order) and isNumeric(Item_Id) and isNumeric(Attribute_Price_difference) and isNumeric(Attribute_Weight_difference) and isNumeric(Use_Item) and isNumeric(IItem_ID) and isNumeric(IItem_Quantity) and isNumeric(Default) then
				if Use_Item<>0 and IItem_Quantity=0 then
                                   fn_error "If choosing, use inventory item for this attribute please indicate a quantity greater then 0"
        			end if
        			if Use_Item<>0 and IItem_ID=0 then
                                   fn_error "If choosing, use inventory item for this attribute please indicate a valid item id"
        			end if
                                Attribute_hidden = checkStringForQ(Request.Form("Attribute_hidden"))
				if Request.Form("Attribute_Class_New") <> "" then
					Attribute_Class = trim(checkStringForQ(Request.Form("Attribute_Class_New")))
				else 
					Attribute_Class = trim(checkStringForQ(Request.Form("Attribute_Class_Pick")))
				end if
				Attribute_value = checkStringForQ(Request.Form("Attribute_value"))
	
				Attribute_sku = checkStringForQ(Request.Form("Attribute_sku"))


				if Default = -1 then
					sql_update = "update store_items_attributes set [default]=0 where Attribute_class='"&Attribute_Class&"' and item_id="&item_id&" and store_id="&store_id
					conn_store.execute sql_update
				end if
	
				if Request.Form("op")<>"edit" then
					sql_insert = "Insert into store_items_attributes (IItem_Quantity, IItem_ID, Use_Item, Store_id, Item_Id, Attribute_class, Attribute_value, Attribute_Price_difference, Attribute_Weight_difference, [Default], Attribute_sku, Attribute_hidden, Display_Order, Attribute_Type) Values ("&IItem_Quantity&", "&IItem_ID&", "&Use_Item&", "&Store_id&", "&Item_Id&",'"&Attribute_class&"', '"&Attribute_value&"', "&Attribute_Price_difference&", "&Attribute_Weight_difference&", "&Default&",'"&Attribute_sku&"', '"&Attribute_hidden&"',"&Display_Order&",'"&Attribute_Type&"')"
					session("sql") = sql_insert
          conn_store.Execute sql_insert
				else
					AttributeId=Request.Form ("Attribute_Id")
					if isNumeric(AttributeId) then
						sqlUpdate="update store_items_attributes set IItem_Quantity="&IItem_Quantity&", IItem_ID="&IItem_ID&", Use_Item="&Use_Item&", Attribute_class='" & Attribute_class & "', " & _
							"Attribute_value='" & Attribute_value & "', " & _	
							"Attribute_Price_difference=" & Attribute_Price_difference & ", " & _			
							"Attribute_Weight_difference=" & Attribute_Weight_difference & ", " & _ 		
							"[Default]=" & Default & ", " & _			
							"Attribute_SKU='" & Attribute_SKU & "', " & _
							"Attribute_Hidden='" & Attribute_Hidden & "', " & _
											  "Display_Order="&Display_Order & " " &_
							"where store_id=" & store_id & " and Attribute_ID=" & AttributeId
	                         
						conn_store.Execute sqlUpdate

						sqlUpdate="update store_items_attributes set Attribute_Type='"&Attribute_Type&"' where store_id="&Store_Id&" and Item_Id="&Item_Id&" and Attribute_class='"&Attribute_class&"'"
						conn_store.Execute sqlUpdate
					else
						Response.Redirect "admin_error.asp?message_id=1"
					end if
				end if
	
				'INSERT / UPDATE ATTRIBUTE
				sql_insert = "update store_items set Sys_Modified='"&now()&"' where Item_Id="&Item_Id&" AND store_id="&store_id
				conn_store.Execute sql_insert

				Response.Redirect "Item_attributes.asp?Item_Id="&Item_Id&"&"&sAddString
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
		'ADD / EDIT ITEM ACCESORIES
		Case "Add_Accessories"
			op=Request.Form ("op")
			if op="" then
				'ADD ITEM ACCESORIES
				Item_Id = Request.Form("Item_Id")
				if isNumeric(Item_Id) then
					'SPLIT TWICE THEN INSERT TO ITEM_ACCESSORIES TABLE
					Items_Array = split(Request.Form("Items_Id"),",")
					for each One_Items_Array in Items_Array
						Accessory_Item_Id = One_Items_Array
						' now insert into table ...
						sql_insert = "insert into Store_Items_Accessories (Store_id,Item_Id,Accessory_Item_Id) Values ("&Store_id&","&Item_Id&","&Accessory_Item_Id&")"
						conn_store.Execute sql_insert
					next
					sql_insert = "update store_items set Sys_Modified='"&now()&"' where Item_Id="&Item_Id&" AND store_id="&store_id
					conn_store.Execute sql_insert
					Response.Redirect "Item_Accessories.asp?Item_Id="&Item_Id&"&"&sAddString
				else
					Response.Redirect "admin_error.asp?message_id=1"
				end if
			else
				'EDIT ITEM ACCESORIES
				AccessoryLineId=Request.Form("Accessory_Line_Id") 
				Item_Id = Request.Form("Item_Id")
				if isNumeric(AccessoryLineId) and isNumeric(Item_Id) then
					Items_Array = split(Request.Form("Items_Id"),",")
					for each One_Items_Array in Items_Array
						Accessory_Item_Id = One_Items_Array
					    sqlUpdate="update Store_items_accessories set Item_Id=" & Item_Id & ", " & _
						    "Accessory_Item_Id=" & Accessory_Item_Id &_	
						    "where store_id=" & store_id & " and Accessory_Line_Id=" & AccessoryLineId
					    conn_store.Execute sqlUpdate
					Next
					Response.Redirect "Item_Accessories.asp?Item_Id="&Item_Id&"&"&sAddString
				else
					Response.Redirect "admin_error.asp?message_id=1"
				end if
			end if
		
		case "Add_Config"
			Item_ID = request.form("Item_ID")
			Config_ID = request.form("Config_ID")
			if isNumeric(Item_ID) and (isNumeric(Config_ID) or Config_ID = "") then
				Config_Name = checkStringForQ(request.form("Config_Name"))
				Config_Desc = checkStringForQ(request.form("Config_Desc"))
				Op = request.form("Op")
				if Op="edit" then
					sql_update = "update store_items_configs set config_name='"&Config_Name&"', config_desc='"&Config_Desc&"' where store_id="&Store_ID&" and config_id="&Config_ID&" and item_id="&Item_ID
				else
					sql_update = "Insert into store_items_configs (Store_ID, Item_ID, Config_Name, Config_Desc) values ("&store_id&", "&Item_ID&", '"&Config_Name&"', '"&Config_Desc&"')"
				end if
				conn_store.execute sql_update
				response.redirect "Item_Configs.asp?Item_ID="&Item_ID&"&"&sAddString
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
		case "Add_Dets"
			Item_ID = request.form("Item_ID")
			config_id = request.form("config_id")
			if isNumeric(Item_ID) and isNumeric(config_id) then
				op = request.form("op")
				Old_Option_Name = checkStringForQ(request.form("Old_Option_Name"))
				Option_Name = checkStringForQ(request.form("Option_Name"))
				if op = "edit" then
					sql_del = "delete from Store_Items_Configs_Dets where store_id="&store_id&" and item_id="&item_id&" and config_id="&config_id&" and name='"&Old_Option_Name&"'"
				else
					sql_del = "delete from Store_Items_Configs_Dets where store_id="&store_id&" and item_id="&item_id&" and config_id="&config_id&" and name='"&Option_Name&"'"
				end if
				conn_store.execute sql_del
				sql_sel_attribs = "select distinct Attribute_class from store_items_attributes where store_id="&store_id&" and item_id="&item_id
				rs_store.open sql_sel_attribs, conn_store, 1, 1
				do while not rs_store.eof
					if request.form(rs_store("Attribute_class"))<>"" then
						sql_insert = "insert into Store_Items_Configs_Dets (Item_ID, Config_ID, Store_ID, Name, Attributes_Class, Attributes_Values) values ("&Item_ID&", "&Config_ID&", "&store_id&", '"&Option_Name&"', '"&checkStringForQ(rs_store("Attribute_class"))&"', '"&checkStringForQ(request.form(rs_store("Attribute_class")))&"')"
						conn_store.execute sql_insert
					else
						rs_store.close
						response.redirect "admin_error.asp?message_id=93"
					end if
					rs_store.movenext
				loop
				rs_store.close
				response.redirect "Item_Configs.asp?Item_ID="&Item_ID&"&"&sAddString
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
			
	end select	  

	'DELETE ITEM ATTRIBUTES
	if Request.QueryString("Delete_Attribute_Id") <> "" then 
		Item_Id = Request.QueryString("Item_Id")
		Attribute_Id = Request.QueryString("Delete_Attribute_Id")
		if isNumeric(Item_Id) and isNumeric(Attribute_Id) then
			sql_delete = "Delete from store_Items_Attributes where Attribute_Id = "&Attribute_Id
			conn_store.Execute sql_delete
			sql_insert = "update store_items set Sys_Modified='"&now()&"' where Item_Id="&Item_Id&" AND store_id="&store_id
			conn_store.Execute sql_insert
			Response.Redirect "Item_attributes.asp?Item_Id="&Item_Id&"&"&sAddString
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if

	'DELETE ITEM ACCESSORIES
	if Request.QueryString("Delete_Accessory_Line_Id") <> "" then 
		Item_Id = Request.QueryString("Item_Id")
		Accessory_Line_Id = Request.QueryString("Delete_Accessory_Line_Id")
		Attribute_Id = Request.QueryString("Delete_Attribute_Id")
		if isNumeric(Item_Id) and isNumeric(Accessory_Line_Id) and isNumeric(Attribute_Id) then
			sql_delete = "Delete from store_Items_Accessories where Accessory_Line_Id = "&Accessory_Line_Id
			conn_store.Execute sql_delete
			sql_insert = "update store_items set Sys_Modified='"&now()&"' where Item_Id="&Item_Id&" AND store_id="&store_id
			conn_store.Execute sql_insert
	
			Response.Redirect "Item_Accessories.asp?Item_Id="&Item_Id&"&"&sAddString
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if
	
	if Request.QueryString("Delete_Config") <> "" then 
		Item_Id = Request.QueryString("Item_Id")
		Config_ID = Request.QueryString("Config_Id")
		if isNumeric(Item_Id) and isNumeric(Config_ID) then
			sql_del = "delete from store_items_configs where store_id="&store_id&" and item_id="&item_id&" and config_id="&config_id
			conn_store.Execute sql_del
			sql_del = "delete from store_items_configs_dets where store_id="&store_id&" and item_id="&item_id&" and config_id="&config_id
			conn_store.Execute sql_del
			response.redirect "Item_Configs.asp?Item_ID="&Item_ID&"&"&sAddString
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if
	
	if Request.QueryString("Delete_Option") <> "" then 
		Item_Id = Request.QueryString("Item_Id")
		Config_ID = Request.QueryString("Config_Id")
		if isNumeric(Item_Id) and isNumeric(Config_ID) then
			sqlConfig="select Name from store_items_configs_dets where config_id="&Config_ID&" and store_id=" & store_id&" and Det_ID="&Request.QueryString("Class")
			rs_Store.open sqlConfig,conn_store,1,1
                        if not rs_store.eof and not rs_store.bof then
			   Option_Name = rs_Store("Name")
			   sql_del = "delete from Store_Items_Configs_Dets where store_id="&store_id&" and item_id="&item_id&" and config_id="&config_id&" and name='"&Option_Name&"'"
			   conn_store.execute sql_del
                        end if
			rs_store.close
			response.redirect "Item_Configs.asp?Item_ID="&Item_ID&"&"&sAddString
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if

end if 

%>
