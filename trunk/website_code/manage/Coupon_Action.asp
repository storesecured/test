<!--#include file="Global_Settings.asp"-->


<%

If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

' ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%> <!--#include virtual="common/Error_Template.asp"--><% 
else

	'CREATE / UPDATE A COUPON
	if Request.Form("Form_Name") = "Create_Coupon" then
		Coupon_Total_Usage = Request.Form("Coupon_Total_Usage")
		Coupon_Total_Left = Request.Form("Coupon_Total_Usage")
		Coupon_Customer_Usage = Request.Form("Coupon_Customer_Usage")
		Total_From=request.form("Total_From")
		Total_To=request.form("Total_To")
      if not isNumeric(Coupon_Total_Usage) and not isNumeric(Coupon_Total_Left) and not isNumeric(Coupon_Customer_Usage) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		Coupon_name = checkStringForQ(Request.Form("Coupon_name"))

		Coupon_Id = checkStringForQ(Request.Form("Coupon_Id"))

		if Request.Form("op")<>"edit" then
			'CHECK IF THE COUPON ID WAS NOT ALLREADY USED
			sql_check="select coupon_id from store_coupons where coupon_id = '"&Coupon_Id&"' and store_id = "&store_id
			rs_store.open sql_check,conn_store
			if rs_store.bof = false then
				rs_store.close
				response.redirect "admin_error.asp?message_id=25"	
			end if
			rs_store.close
		end if

		' RETRIEVING FORM DATA
		Coupon_Start_Date = checkStringForQ(Request.Form("Coupon_Start_Date"))
		Coupon_End_Date = checkStringForQ(Request.Form("Coupon_End_Date"))
		Coupon_Type = Request.Form("Coupon_Type")
		if Coupon_Type = 1 then
			Coupon_Amount = Request.Form("Coupon_Fixed_Amount")
		else
			Coupon_Amount = Request.Form("Coupon_Percent_of_Sale")
		end if
		
		if not isNumeric(Coupon_Amount) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		
		if Request.Form("Discount_Items") = "All" then
				Discounted_Items_Skus = "All"
				Discounted_Items_Ids=""
				Is_Exclusion=0
		else
			  	if Request.Form("Discount_Items") = "part" then
			  	  Is_Exclusion=0
			  	else
			  	  Is_Exclusion=-1
                                end if
                                Discounted_Items_Skus = replace(checkStringForQ(Request.Form("Discounted_Items_Skus")),", ",",")
		                Discounted_Items_Skus = replace(Discounted_Items_Skus,",","','")
                                sql_select = "select item_id from store_items WITH (NOLOCK) where store_id="&store_id&" and item_sku in ('"&Discounted_Items_Skus&"')"

                                set myfields=server.createobject("scripting.dictionary")
	                        Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	                        if noRecords = 0 then
	                        FOR rowcounter= 0 TO myfields("rowcount")
	                               if Discounted_Items_Ids="" then
	                                       Discounted_Items_Ids =mydata(myfields("item_id"),rowcounter)
	                               else
	                                       Discounted_Items_Ids =Discounted_Items_Ids&","&mydata(myfields("item_id"),rowcounter)
	                               end if
	                        Next
	                        end if
                end if


		' INSERTING / UPDATING COUPON IN STORE_COUPONS TABLE
		if Request.Form("op")<>"edit" then
			' now we are ready to insert to table ...
			sql_insert = "insert into store_Coupons (Store_id,Coupon_name,Coupon_Start_Date,Coupon_End_Date,Coupon_Id,Coupon_Type,Coupon_Amount,Coupon_Total_Usage,Coupon_Total_Left,Coupon_Customer_Usage,Discounted_Items_Ids,Total_From,Total_To,Is_Exclusion) values ("&Store_id&",'"&Coupon_name&"','"&Coupon_Start_Date&"','"&Coupon_End_Date&"','"&Coupon_Id&"',"&Coupon_Type&","&Coupon_Amount&","&Coupon_Total_Usage&","&Coupon_Total_Left&","&Coupon_Customer_Usage&",'"&Discounted_Items_Ids&"',"&Total_From&","&Total_To&","&Is_Exclusion&")"
			session("sql") = sql_insert
         conn_store.Execute sql_insert
		else
			CouponLineId=Request.Form ("Coupon_Line_Id")
			Total_From=request.form("Total_From")
			Total_To=request.form("Total_To")

			if not isNumeric(CouponLineId) then
				Response.Redirect "admin_error.asp?message_id=1"
			end if

			' CHECK IF THE COUPON ID IS NOT ALLREADY USED
			sql_check="select coupon_id from store_coupons where coupon_id = '"&Coupon_Id&"' and store_id = "&store_id & _
				" and coupon_line_id<>" & CouponLineId
			rs_store.open sql_check,conn_store

			if rs_store.bof = false then
				rs_store.close
				response.redirect "admin_error.asp?message_id=25"	
			end if
			rs_store.close	
			sqlUpdate="update store_coupons set Coupon_name='" & Coupon_name & "', " & _
					"Coupon_Start_Date='" & Coupon_Start_Date & "', " & _
					"Coupon_End_Date='" & Coupon_End_Date & "', " & _
					"Coupon_Id='" & Coupon_Id & "', " & _	
					"Coupon_Type=" & Coupon_Type & ", " & _ 
					"Coupon_Amount=" & Coupon_Amount & ", " & _	
					"Coupon_Total_Usage=" & Coupon_Total_Usage & ", " & _	
					"Coupon_Customer_Usage=" & Coupon_Customer_Usage & ", " & _
                                        "Discounted_Items_Ids='" & Discounted_Items_Ids & "', " & _
                                        "Total_From=" & Total_From & ", " & _
					"Total_To=" & Total_To & ", " & _
					"Is_Exclusion=" & Is_Exclusion & " " & _
					"where store_id=" & store_id & " and coupon_line_id=" & CouponLineId

                        


      Session("sql") = sqlUpdate
			conn_store.Execute sqlUpdate
		end if
	
		' UPDATING STORE_SETTINGS - SET ENABLE_COUPON=TRUE
		sql_update = "update store_settings set Enable_Coupon = -1 where Store_id = "&Store_id
		conn_store.Execute sql_update		
		Response.Redirect "Coupon_manager.asp"
	end if 

	' DELETE A COUPON
	if Request.QueryString("Delete_Coupon_Line_Id") <> "" then 
		Coupon_Line_Id = Request.QueryString("Delete_Coupon_Line_Id")
		if not isNumeric(Coupon_Line_Id) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		sql_delete = "Delete from store_Coupons where Coupon_Line_Id = "&Coupon_Line_Id
		conn_store.Execute sql_delete
		' IF THERE ARE NO OTHER COUPONS IN THE DATABASE
		sql_check = "select coupon_id from store_Coupons where Store_id = "&Store_id
		rs_store.open sql_check,conn_store
		if rs_store.bof = true then
			' UPDATING STORE_SETTINGS - SET ENABLE_COUPON=FALSE
			sql_update = "update store_settings set Enable_Coupon = 0 where Store_id = "&Store_id
			conn_store.Execute sql_update
		end if	
		Response.Redirect "Coupon_manager.asp"
	end if

End If 

%>
