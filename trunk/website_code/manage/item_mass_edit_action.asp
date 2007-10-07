<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
'USER CLICKED ON "MASS EDIT"
'UPDATE ALL SELECTED ITEMS ACCORDING TO THE MODIFICATIONS MADE BY USER
if request.form("Form_Name")= "Mass_Edit_Update" then

If Form_Error_Handler(Request.Form) <> "" then
   Error_Log = Form_Error_Handler(Request.Form)
   if Error_Log <> "" then
	%> <!--#include virtual="common/Error_Template.asp"--><%
	response.end
	end if
end if

	sItem_ids = Request.Form("DELETE_IDS")
	updates_str = ""
	if request.form("Description_S")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Description_S='"&nullifyQ(request.form("Description_S"))&"' "
	end if


	
	
	if request.form("Retail_Price")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Retail_Price="&request.form("Retail_Price")&" "
	end if

	if request.form("Wholesale_Price")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Wholesale_Price="&request.form("Wholesale_Price")&" "
	end if



	if request.form("Use_Price_By_Matrix")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Use_Price_By_Matrix="&request.form("Use_Price_By_Matrix")&" "
	end if

	if request.form("Show_Homepage")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Show_Homepage="&request.form("Show_Homepage")&" "
	end if



	if request.form("Hide_Price")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Hide_Price="&request.form("Hide_Price")&" "
	end if


	if request.form("Cust_Price")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Enable_Cust_Price="&request.form("Cust_Price")&" "
	end if



	if request.form("Description_L")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Description_L='"&nullifyQ(request.form("Description_L"))&"' "
	end if



	if request.form("ImageS_Path")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " ImageS_Path='"&request.form("ImageS_Path")&"' "
	end if

	if request.form("ImageL_Path")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " ImageL_Path='"&request.form("ImageL_Path")&"' "
	end if




	if request.form("Item_Handling")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Item_Handling="&request.form("Item_Handling")&" "
	end if



	if request.form("Waive_Shipping")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Waive_Shipping="&request.form("Waive_Shipping")&" "
	end if


	if request.form("Ship_Location_Id")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Ship_Location_Id="&request.form("Ship_Location_Id")&" "
	end if





	if request.form("Quantity_in_stock")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Quantity_in_stock="&request.form("Quantity_in_stock")&" "
	end if
	if request.form("Item_Weight")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Item_Weight="&request.form("Item_Weight")&" "
	end if
	if request.form("Shipping_Fee")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Shipping_Fee="&request.form("Shipping_Fee")&" "
	end if
	
	if request.form("Retail_Price_special_Discount")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Retail_Price_special_Discount="&request.form("Retail_Price_special_Discount")&" "
	end if
	if request.form("Special_Start_Date")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Special_Start_Date='"&request.form("Special_Start_Date")&"' "
	end if
	if request.form("Special_End_Date")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Special_End_Date='"&request.form("Special_End_Date")&"' "
	end if
	
	if request.form("Taxable")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Taxable="&request.form("Taxable")&" "
	end if
	if request.form("Show")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Show="&request.form("Show")&" "
	end if
	if request.form("Fractional")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Fractional="&request.form("Fractional")&" "
	end if



	if request.form("Quantity_Control")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Quantity_Control="&request.form("Quantity_Control")&" "
	end if


	if request.form("Quantity_Control_Number")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Quantity_Control_Number="&request.form("Quantity_Control_Number")&" "
	end if



	if request.form("Hide_Stock")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Hide_Stock="&request.form("Hide_Stock")&" "
	end if

	if request.form("Quantity_Minimum")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Quantity_Minimum="&request.form("Quantity_Minimum")&" "
	end if



	if request.form("Meta_Title")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Meta_Title='"&checkStringForQ(request.form("Meta_Title"))&"' "
	end if

	if request.form("Meta_Keywords")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Meta_Keywords='"&checkStringForQ(request.form("Meta_Keywords"))&"' "
	end if

	if request.form("Meta_Description")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Meta_Description='"&checkStringForQ(request.form("Meta_Description"))&"' "
	end if



	if request.form("Recurring_fee")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Recurring_fee="&request.form("Recurring_fee")&" "
	end if


	if request.form("Recurring_days")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Recurring_days="&request.form("Recurring_days")&" "
	end if
	if request.form("Brand")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Brand='"&checkStringForQ(request.form("Brand"))&"' "
	end if
	if request.form("Condition")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Condition='"&checkStringForQ(request.form("Condition"))&"' "
	end if
	if request.form("Product_Type")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Product_Type='"&checkStringForQ(request.form("Product_Type"))&"' "
	end if

	if request.form("File_Location")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " File_Location='"&request.form("File_Location")&"' "
	end if

	if request.form("Item_pin")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Item_pin="&request.form("Item_pin")&" "
	end if


	if request.form("Item_Remarks")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " Item_Remarks='"&request.form("Item_Remarks")&"' "
	end if


	if request.form("U_d_1")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_1="&nullifyQ(request.form("U_d_1"))&" "
	end if

	if request.form("U_d_1_name")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_1_name='"&request.form("U_d_1_name")&"' "
	end if

	if request.form("U_d_2")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_2="&nullifyQ(request.form("U_d_2"))&" "
	end if

	if request.form("U_d_2_name")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_2_name='"&request.form("U_d_2_name")&"' "
	end if



	if request.form("U_d_3")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_3="&nullifyQ(request.form("U_d_3"))&" "
	end if

	if request.form("U_d_3_name")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_3_name='"&request.form("U_d_3_name")&"' "
	end if


	if request.form("U_d_4") <>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_4="&nullifyQ(request.form("U_d_4"))&" "
	end if

	if request.form("U_d_4_name")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_4_name='"&request.form("U_d_4_name")&"' "
	end if

	if request.form("U_d_5")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_5="&nullifyQ(request.form("U_d_5"))&" "
	end if

	if request.form("U_d_5_name")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " U_d_5_name='"&request.form("U_d_5_name")&"' "
	end if


	if request.form("M_d_1")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " M_d_1='"&nullifyQ(request.form("M_d_1"))&"' "
	end if

	if request.form("M_d_2")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " M_d_2='"&nullifyQ(request.form("M_d_2"))&"' "
	end if

	if request.form("M_d_3")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " M_d_3='"&nullifyQ(request.form("M_d_3"))&"' "
	end if
	if request.form("M_d_4")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " M_d_4='"&nullifyQ(request.form("M_d_4"))&"' "
	end if
	if request.form("M_d_5")<>"" then
		if updates_str<>"" then
			updates_str = updates_str&" , "
		end if
		updates_str = updates_str & " M_d_5='"&nullifyQ(request.form("M_d_5"))&"' "
	end if

     if request.form("Department_ID")<>"" then
          updates_str = updates_str&" "

		for each sSingleItem in split(sItem_ids,",")
	     	sql_update = "sp_item_dept_update "&store_id&","&sSingleItem&",'"&request.form("Department_ID")&"'"
			conn_store.execute(sql_update)
          next
	end if


	if updates_str<>"" then
		updates_str = "Update Store_Items set sys_modified = '"&GetUSFormatDate(now)&"', "&updates_str&" where item_id in ("&sItem_ids&") and store_id="&store_id
		
		
    conn_store.Execute updates_str
	end if

end if
response.redirect "edit_items.asp"
%>