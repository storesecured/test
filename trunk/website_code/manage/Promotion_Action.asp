<!--#include file="Global_Settings.asp"-->
<%
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><% 
	
else
	
	'CREATE / EDIT PROMOTION
	if Request.Form("Form_Name") = "Create_Promotion" then
		'RETRIEVE FORM DATA
		Discounted_items_want = Request.Form("Discounted_items_want")
		Customer_group = Request.Form("Customer_group")
		Total_From = Request.Form("Total_From")
		Total_To = Request.Form("Total_To")
		if not isNumeric(Discounted_Items_Amount) or not isNumeric(Customer_group) or not isNumeric(Total_From) or not isNumeric(Total_To) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if

		Promotion_name = replace(Request.Form("Promotion_name"),"'","''")
		Promotion_Start_Date = checkStringForQ(Request.Form("Promotion_Start_Date"))
		Promotion_End_Date = checkStringForQ(Request.Form("Promotion_End_Date"))
                Is_Exclusion=0
                
		if Discounted_items_want = 1 then
			Discounted_Items_Amount = Request.Form("Discounted_Items_Amount")
			if not isNumeric(Discounted_Items_Amount) then
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
		else
			Discounted_Items_Amount = 0
			Discounted_Items_Skus = ""
			Discounted_Items_Ids=""
		end if
		Free_items_want = Request.Form("Free_items_want")
		if Free_items_want = 1 then 
			Free_items = checkStringForQ(Request.Form("Free_items"))
			Free_items = replace(checkStringForQ(Free_items),", ",",")
		        Free_items = replace(Free_items,",","','")
                        sql_select = "select item_id from store_items WITH (NOLOCK) where store_id="&store_id&" and item_sku in ('"&Free_items&"')"

                        set myfields=server.createobject("scripting.dictionary")
	                Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	                if noRecords = 0 then
	                FOR rowcounter= 0 TO myfields("rowcount")
	                        if Free_items_Ids="" then
	                                Free_items_Ids =mydata(myfields("item_id"),rowcounter)
	                        else
	                                Free_items_Ids =Free_items_Ids&","&mydata(myfields("item_id"),rowcounter)
	                        end if
	                Next
	                end if

		else
			Free_items = ""
		end if


		if Request.Form("op")<>"edit" then
			
			'IF ADD, THEN PREPARE AND RUN AN INSERT
			sql_insert = "insert into store_promotions_rules (Store_id,Promotion_name,Promotion_Start_Date,Promotion_End_Date,Discounted_Items_Amount,Discounted_Items_Ids,Free_items_ids,Total_From,Total_To,Customer_group,Is_Exclusion) values ("&Store_id&",'"&Promotion_name&"','"&Promotion_Start_Date&"','"&Promotion_End_Date&"',"&Discounted_Items_Amount&",'"&Discounted_Items_Ids&"','"&Free_items_Ids&"',"&Total_From&","&Total_To&","&Customer_group&","&Is_Exclusion&")"
			conn_store.Execute sql_insert
			
			' added 05/07/2005 [mm/dd/yy]
			' get all the promotion_ids for the store
			
			sql_promotion_id = "select promotion_id from store_promotions_rules where store_id = "&Store_Id
			rs_store.open sql_promotion_id,conn_store,1,1
			strPromotionId = ""
			
			if not rs_store.eof then
				while not rs_store.eof
				intStoreId = rs_store(0)
				strPromotionId = strPromotionId&","&intStoreId
				rs_store.movenext
				wend
			strPromotionId = right(strPromotionId,len(strPromotionId)-1)
			end if			
			
			rs_store.close()
			
			
			response.redirect "Promotion_manager.asp"
			
		else
			'IF EDIT, THEN PREPARE AND RUN AN UPDATE
			PromotionId=Request.Form ("Promotion_Id")
			sqlUpdate="update store_promotions_rules set Promotion_name='" & Promotion_name & "', " & _
					"Promotion_Start_Date='" & Promotion_Start_Date & "', " & _
					"Promotion_End_Date='" & Promotion_End_Date & "', " & _
					"Discounted_Items_Amount=" & Discounted_Items_Amount & ", " & _ 
					"Discounted_Items_Ids='" & Discounted_Items_Ids & "', " & _
					"Free_items_Ids='" & Free_items_Ids & "', " & _ 
					"Total_From=" & Total_From & ", " & _	
					"Total_To=" & Total_To & ", " & _
                                        "Is_Exclusion=" & Is_Exclusion & ", " & _
					"Customer_group='" & Customer_group & "' " & _
					"where store_id=" & store_id & " and promotion_id=" & PromotionId

                        session("sql") = sqlUpdate
         conn_store.Execute sqlUpdate
			
                
			response.redirect "Promotion_manager.asp"
	          
		end if
	end if

	'UPDATE AVAILABLE PROMOTIONS
	if Request.Form("Form_Name") = "update_promotion_ids" then
		'UPDATE THE STORE_SETTINGS TABLE WITH THE NEW PROMOTIONS
		Promotion_Id =Request.Form("Promotion_Id")
		if not isNumeric(Promotion_Id) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		response.redirect "Promotion_manager.asp"
	     
	end if

	'DELETE PROMOTION
	if Request.QueryString("Delete_Promotion_Id") <> "" then 
		Promotion_Id = Request.QueryString("Delete_Promotion_Id")
		if not isNumeric(Promotion_Id) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		sql_delete = "delete from store_promotions_rules where Promotion_Id = "&Promotion_Id
		conn_store.Execute sql_delete
		response.redirect "Promotion_manager.asp"

	end if

End If

%> 
