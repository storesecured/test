<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
'USER CLICKED ON "QUICK EDIT"
'UPDATE EACH OF SELECTED ITEMS ACCORDING TO THE MODIFICATIONS MADE BY USER

if request.form("Form_Name")= "Quick_Edit_Update" then

	If Form_Error_Handler(Request.Form) <> "" then 
		Error_Log = Form_Error_Handler(Request.Form)
		if Error_Log <> "" then
      %> <!--#include virtual="common/Error_Template.asp"--><%
      response.end
		end if

	end if

	buffer = Request.Form("DELETE_IDS")
	pos = 0
	
	do while pos<>-1
		pos = instr(buffer,",")
		if pos<>0 then
			currentItem = mid(buffer,1,pos-1)
			buffer = mid (buffer,pos+1)
		else
			currentItem = buffer
			pos = -1
		end if
		if currentItem<>"" then
			do while mid(currentItem,1,1)=" "
				currentItem = mid(currentItem,2)
			loop

			updates_str = ""
			
			updates_str = updates_str & ", Item_Sku='"&checkStringForQ(request.form("Item_Sku_"&currentItem))&"' "
			updates_str = updates_str & ", Item_Name='"&checkStringForQ(request.form("Item_Name_"&currentItem))&"' "
			updates_str = updates_str & ", Item_Page_Name='"&checkStringForQ(request.form("Item_Page_Name_"&currentItem))&"' "


			updates_str = updates_str & ", Retail_Price="&request.form("Retail_Price_"&currentItem)&" "

			updates_str = updates_str & ", Wholesale_Price="&request.form("Wholesale_Price_"&currentItem)&" "
			
			
			if request.form("Use_Price_By_Matrix_"&currentItem)<>"" then
				updates_str = updates_str & ", Use_Price_By_Matrix=-1 "
			else
				updates_str = updates_str & ", Use_Price_By_Matrix=0 "
			end if


			if request.form("Show_Homepage_"&currentItem)<>"" then
				updates_str = updates_str & ", Show_Homepage=-1 "
			else
				updates_str = updates_str & ", Show_Homepage=0 "
			end if
			

			if request.form("Taxable_"&currentItem)<>"" then
				updates_str = updates_str & ", Taxable=-1 "
			else
				updates_str = updates_str & ", Taxable=0 "
			end if



			if request.form("Show_"&currentItem)<>"" then
				updates_str = updates_str & ", Show=-1 "
			else
				updates_str = updates_str & ", Show=0 "
			end if



			if request.form("Hide_Price_"&currentItem)<>"" then
				updates_str = updates_str & ", Hide_Price=-1 "
			else
				updates_str = updates_str & ", Hide_Price=0 "
			end if

			
			if request.form("Cust_Price_"&currentItem)<>"" then
				updates_str = updates_str & ", enable_cust_price=-1 "
			else
				updates_str = updates_str & ", enable_cust_price=0 "
			end if


			
			updates_str = updates_str & ", View_Order="&request.form("View_Order_"&currentItem)&" "


			updates_str = updates_str & ", Description_S='"&nullifyQ(request.form("Description_S_"&currentItem))&"' "
			updates_str = updates_str & ", ImageS_path='"&checkStringForQ(request.form("ImageS_path_"&currentItem))&"' "
			updates_str = updates_str & ", Description_L='"&nullifyQ(request.form("Description_L_"&currentItem))&"' "
			updates_str = updates_str & ", ImageL_Path='"&checkStringForQ(request.form("ImageL_Path_"&currentItem))&"' "



			updates_str = updates_str & ", Item_Weight="&request.form("Item_Weight_"&currentItem)&" "
			updates_str = updates_str & ", Item_Handling="&request.form("Item_Handling_"&currentItem)&" "



			if request.form("Waive_Shipping_"&currentItem)<>"" then
				updates_str = updates_str & ", Waive_Shipping=-1 "
			else
				updates_str = updates_str & ", Waive_Shipping=0 "
			end if


			updates_str = updates_str & ", Ship_Location_Id="&request.form("Ship_Location_Id_"&currentItem)&" "

			updates_str = updates_str & ", Shipping_Fee="&request.form("Shipping_Fee_"&currentItem)&" "


			updates_str = updates_str & ", Quantity_in_stock="&request.form("Quantity_in_stock_"&currentItem)&" "

			if request.form("Quantity_Control_"&currentItem)<>"" then
				updates_str = updates_str & ", Quantity_Control=-1 "
			else
				updates_str = updates_str & ", Quantity_Control=0 "
			end if


			if request.form("Hide_Stock_"&currentItem)<>"" then
				updates_str = updates_str & ", Hide_Stock=-1 "
			else
				updates_str = updates_str & ", Hide_Stock=0 "
			end if




			updates_str = updates_str & ", Quantity_Control_Number="&request.form("Quantity_Control_Number_"&currentItem)&" "
			updates_str = updates_str & ", Quantity_Minimum="&request.form("Quantity_Minimum_"&currentItem)&" "

			if request.form("Fractional_"&currentItem)<>"" then
				updates_str = updates_str & ", Fractional=-1 "
			else
				updates_str = updates_str & ", Fractional=0 "
			end if


			
			updates_str = updates_str & ", Meta_Title='"&checkStringForQ(request.form("Meta_Title_"&currentItem))&"' "
			updates_str = updates_str & ", Meta_Keywords='"&checkStringForQ(request.form("Meta_Keywords_"&currentItem))&"' "
			updates_str = updates_str & ", Meta_Description='"&checkStringForQ(request.form("Meta_Description_"&currentItem))&"' "
			
			updates_str = updates_str & ", Recurring_days="&request.form("Recurring_days_"&currentItem)&" "
			updates_str = updates_str & ", Recurring_fee="&request.form("Recurring_fee_"&currentItem)&" "
			updates_str = updates_str & ", Brand='"&checkStringForQ(request.form("Brand_"&currentItem))&"' "
			updates_str = updates_str & ", Condition='"&checkStringForQ(request.form("Condition_"&currentItem))&"' "
			updates_str = updates_str & ", Product_type='"&checkStringForQ(request.form("Product_type_"&currentItem))&"' "
			updates_str = updates_str & ", Retail_Price_Special_Discount="&request.form("Retail_Price_Special_Discount_"&currentItem)&" "
			updates_str = updates_str & ", Special_Start_Date='"&request.form("Special_Start_Date_"&currentItem)&"' "
			updates_str = updates_str & ", Special_End_Date='"&request.form("Special_End_Date_"&currentItem)&"' "
			
		
			updates_str = updates_str & ", File_Location='"&checkStringForQ(request.form("File_Location_"&currentItem))&"' "
		
			if request.form("Item_pin_"&currentItem)<>"" then
				updates_str = updates_str & ", Item_pin=-1 "
			else
				updates_str = updates_str & ", Item_pin=0 "
			end if

	
			
			
			updates_str = updates_str & ", Item_Remarks='"&checkStringForQ(request.form("Item_Remarks_"&currentItem))&"' "
			
			if request.form("U_d_1_"&currentItem)<>"" then
				updates_str = updates_str & ", U_d_1=-1 "
			else
				updates_str = updates_str & ", U_d_1=0 "
			end if
			updates_str = updates_str & ", U_d_1_name='"&checkStringForQ(request.form("U_d_1_name_"&currentItem))&"' "

			if request.form("U_d_2_"&currentItem)<>"" then
				updates_str = updates_str & ", U_d_2=-1 "
			else
				updates_str = updates_str & ", U_d_2=0 "
			end if
			updates_str = updates_str & ", U_d_2_name='"&checkStringForQ(request.form("U_d_2_name_"&currentItem))&"' "


			if request.form("U_d_3_"&currentItem)<>"" then
				updates_str = updates_str & ", U_d_3=-1 "
			else
				updates_str = updates_str & ", U_d_3=0 "
			end if
			updates_str = updates_str & ", U_d_3_name='"&checkStringForQ(request.form("U_d_3_name_"&currentItem))&"' "


			if request.form("U_d_4_"&currentItem)<>"" then
				updates_str = updates_str & ", U_d_4=-1 "
			else
				updates_str = updates_str & ", U_d_4=0 "
			end if
			updates_str = updates_str & ", U_d_4_name='"&checkStringForQ(request.form("U_d_4_name_"&currentItem))&"' "


			if request.form("U_d_5_"&currentItem)<>"" then
				updates_str = updates_str & ", U_d_5=-1 "
			else
				updates_str = updates_str & ", U_d_5=0 "
			end if

			updates_str = updates_str & ", U_d_5_name='"&checkStringForQ(request.form("U_d_5_name_"&currentItem))&"' "

			updates_str = updates_str & ", M_d_1='"&nullifyQ(request.form("M_d_1_"&currentItem))&"' "
			updates_str = updates_str & ", M_d_2='"&nullifyQ(request.form("M_d_2_"&currentItem))&"' "
			updates_str = updates_str & ", M_d_3='"&nullifyQ(request.form("M_d_3_"&currentItem))&"' "
			updates_str = updates_str & ", M_d_4='"&nullifyQ(request.form("M_d_4_"&currentItem))&"' "
			updates_str = updates_str & ", M_d_5='"&nullifyQ(request.form("M_d_5_"&currentItem))&"' "

			updates_str = updates_str & ", Custom_Link='"&nullifyQ(request.form("Custom_Link_"&currentItem))&"' "

			if replace(request.form("Department_ID_"&currentItem)," ","")<>replace(request.form("Old_Department_ID_"&currentItem)," ","") then
	     		sql_update = "sp_item_dept_update "&store_id&","&currentItem&",'"&request.form("Department_ID_"&currentItem)&"'"
				conn_store.execute(sql_update)
				updates_str = updates_str & " "
			end if


			updates_str = "Update Store_Items set sys_modified = '"&GetUSFormatDate(now)&"' "&updates_str&" where item_id="&currentItem&" and store_id="&store_id

			conn_store.Execute updates_str
			
		end if
	loop
	
end if
response.redirect "edit_items.asp"

%>