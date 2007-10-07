<!--#include file="Global_Settings.asp"-->

<%

'ERROR CHECKING
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if


If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><%
else
	on error goto 0

  'ADD NEW CUSTOMER
	if Request.Form("Add_New_Customer") <> "" then
		Record_type =Request.Form("Record_type")
		User_ID = checkStringForQ(Request.Form("User_ID"))
		Password = checkStringForQ(Request.Form("Password"))
		Last_name = checkStringForQ(Request.Form("Last_name"))
		First_name = checkStringForQ(Request.Form("First_name"))
		Company = checkStringForQ(Request.Form("Company"))
		Address1 = checkStringForQ(Request.Form("Address1"))
		Address2 = checkStringForQ(Request.Form("Address2"))
		City = checkStringForQ(Request.Form("City"))
		State = checkStringForQ(Request.Form("State"))
		Country = checkStringForQ(Request.Form("Country"))
		Phone = checkStringForQ(Request.Form("Phone"))
		EMail = checkStringForQ(Request.Form("EMail"))
		FAX = checkStringForQ(Request.Form("FAX"))
		Zip = checkStringForQ(Request.Form("Zip"))
		Registration_Date = Now ()
		if Request.Form("Spam") <>  "" then
			Spam = Request.Form("Spam")
		else
			Spam = 0
		end if
		if Request.Form("Budget_orig") <> "" then
			Budget_orig = Request.Form("Budget_orig")
			Budget_left = Request.Form("Budget_orig")
		else
			Budget_orig = 0
			Budget_left = 0
		end if
		if Request.Form("Reward_Total") <> "" then
			Reward_Total = Request.Form("Reward_Total")
			Reward_Left = Request.Form("Reward_Total")
		else
			Reward_Total = 0
			Reward_Left = 0
		end if
		if Request.Form("Tax_Exempt") <>  "" then
			Tax_Exempt = Request.Form("Tax_Exempt")
		else
			Tax_Exempt = 0
		end if
		if Request.Form("Protected_Access") <>  "" then
			Protected_Access = Request.Form("Protected_Access")
		else
			Protected_Access = 0
		end if
		Orders_Total = 0	
		Shipping_Same = 1
		if Request.Form("Password") <> Request.Form("Password_Confirm") then
			response.redirect "admin_error.asp?Message_id=10"
		end if
		if Request.Form("C_One_Click_Shopping") <>  "" then
			C_One_Click_Shopping = Request.Form("C_One_Click_Shopping")
		else
			C_One_Click_Shopping = 0
		end if

		if isNull(StartCID) or StartCID="" then
		   StartCID=0
		end if
		
    dim AddCustomerFlag
    AddCustomerFlag = false
    
	if Country="Canada" or Country="United States" then
	    if len(State)>2 then
	        AddCustomerFlag = true
	    else
	        AddCustomerFlag = false
	    end if    
	end if
	
	if AddCustomerFlag = true then
        Response.Redirect "admin_error.asp?Message_Add=In Case of United States and Canada, You can enter only 2 letters in STATE field"	    
	else
		sql_create_customer = "store_register_action "&Store_id&",'"&User_ID&"','"&Password&"','"&Last_name&"','"&First_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Spam&","&StartCID&","&Budget_orig&","&Budget_left&","&Reward_Total&","&Reward_left&","&Protected_Access&","&Tax_Exempt&","&C_One_Click_Shopping
		rs_store.open sql_create_customer,conn_store,1,1
		cid=rs_store("cid")
		ccid=rs_store("ccid")
		rs_Store.close

		if cid=-1 then
		  response.redirect Site_Name&"error.asp?Message_id=101&Message_Add="&Server.urlencode(ccid)
		end if

		Response.Redirect "Modify_customer.asp?cid="&cid&"&ssadr="&Record_type

		cid = int(Request.QueryString("cid") )
	end if
end if	

	'MODIFY BILLING ADDRESS
	if Request.Form("Modify_my_0") <> "" then
		'RETRIEVE FORM DATA
		cid=request.querystring("Cid")
        Record_type = Request.Form("Record_type")
		Last_name = checkStringForQ(Request.Form("Last_name"))
		First_name = checkStringForQ(Request.Form("First_name"))
		Company = checkStringForQ(Request.Form("Company"))
		Address1 = checkStringForQ(Request.Form("Address1"))
		Address2 = checkStringForQ(Request.Form("Address2"))
		City = checkStringForQ(Request.Form("City"))
		State = checkStringForQ(Request.Form("State"))
		Zip = checkStringForQ(Request.Form("Zip"))
		Country = checkStringForQ(Request.Form("Country"))
		Phone = checkStringForQ(Request.Form("Phone"))
		Email = checkStringForQ(Request.Form("Email"))
		Fax = checkStringForQ(Request.Form("Fax"))
		if request.form("Is_Residential")<>"" then
			   Is_Residential=request.form("Is_Residential")
		else
			   Is_Residential=0
		end if

    dim flag
    flag = false
    
	if Country="Canada" or Country="United States" then
	    if len(State)>2 then
	        flag = true
	    else
	        flag = false
	    end if    
	end if
	
	if flag = true then
        Response.Redirect "admin_error.asp?Message_Add=In Case of United States and Canada, You can enter only 2 letters in STATE field"	    
	else
		sql_update_customer = "store_register_update "&Cid&","&Store_id&","&Record_type&",'"&User_ID&"','"&Password&"','"&Last_name&"','"&First_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Is_Residential
		conn_store.Execute sql_update_customer

		if Record_type<>0 then
			Response.Redirect "Modify_customer.asp?cid="&cid&"&ssadr="&Record_type
		else
			Response.Redirect "Modify_customer.asp?cid="&cid
		end if
	end if
	
	end if

	'MODIFY SHIPPING ADDRESSES
	if Request.Form("Modify_my_1") <> "" then
		'RETRIEVE FORM DATA
		cid=request.querystring("Cid")
        Record_type = Request.Form("Record_type")
		Last_name = checkStringForQ(Request.Form("Last_name1"))
		First_name = checkStringForQ(Request.Form("First_name1"))
		Company = checkStringForQ(Request.Form("Company"))
		Address1 = checkStringForQ(Request.Form("Address11"))
		Address2 = checkStringForQ(Request.Form("Address2"))
		City = checkStringForQ(Request.Form("City1"))
		State = checkStringForQ(Request.Form("State1"))
		Country = checkStringForQ(Request.Form("Country1"))
		Phone = checkStringForQ(Request.Form("Phone1"))
		Email = checkStringForQ(Request.Form("Email1"))
		Fax = checkStringForQ(Request.Form("Fax"))
		Zip = checkStringForQ(Request.Form("Zip1"))
		if request.form("Is_Residential")<>"" then
			   Is_Residential=request.form("Is_Residential")
		else
			   Is_Residential=0
		end if

		
	dim ShippingFlag
    ShippingFlag = false
    
	if Country1="Canada" or Country1="United States" then
	    if len(State1)>2 then
	        ShippingFlag = true
	    else
	        ShippingFlag = false
	    end if    
	end if
	
	if ShippingFlag = true then
        Response.Redirect "admin_error.asp?Message_Add=In Case of United States and Canada, You can enter only 2 letters in STATE field"	    
	else
		sql_update_customer = "store_register_update "&Cid&","&Store_id&","&Record_type&",'"&User_ID&"','"&Password&"','"&Last_name&"','"&First_name&"','"&Company&"','"&Address1&"','"&Address2&"','"&City&"','"&Zip&"','"&State&"','"&Country&"','"&Phone&"','"&EMail&"','"&FAX&"',"&Is_Residential
		conn_store.Execute sql_update_customer
		
		if Record_type<>0 then
			Response.Redirect "Modify_customer.asp?cid="&cid&"&ssadr="&Record_type
		else
			Response.Redirect "Modify_customer.asp?cid="&cid
		end if
	end if
end if	

	'MODOFY LOGIN ID & PASSWORD
	if Request.Form("Modify_my_Account") <> "" then
		cid=request.querystring("Cid")
		'RETRIEVE FORM DATA
		Record_type = Request.Form("Record_type")
		User_id = Replace(Request.Form("User_id"), "'", "''")
		Tax_Id = Replace(Request.Form("Tax_Id"), "'", "''")
		Password = Replace(Request.Form("Password"), "'", "''")
		Password_Confirm = Replace(Request.Form("Password_Confirm"), "'", "''")
		budget_orig = Request.Form("budget_orig")
		budget_left = Request.Form("budget_left")
		reward_total = Request.Form("reward_total")
		Reward_Left = Request.Form("Reward_Left")
		if Request.Form("Password") <> Request.Form("Password_Confirm") then
			response.redirect "admin_error.asp?Message_id=10"
		end if
  
		if Request.Form("Spam") <>  "" then
			Spam = Request.Form("Spam")
		else
			Spam = 0
		end if
		if Request.Form("C_One_Click_Shopping") <>  "" then
			C_One_Click_Shopping = Request.Form("C_One_Click_Shopping")
		else
			C_One_Click_Shopping = 0	
		end if
		if Request.Form("Tax_Exempt") <>  "" then
			Tax_Exempt = Request.Form("Tax_Exempt")
		else
			Tax_Exempt = 0  
		end if
		if Request.Form("Protected_Access") <>  "" then
			Protected_Access = -1
		else
			Protected_Access = 0
		end if
		if reward_total = "" then
		   reward_total = 0
		end if
		if reward_left = "" then
		   reward_left = 0
		end if


		' Customer Added Groups
		if Request.Form("Group_id") <>  "" then
			Additional_Groups = checkStringForQ(Request.Form("Group_id"))
		else
			Additional_Groups = null
		end if
		
		CustId = checkStringForQ(Request.Form("CCid"))
		

		' Get Records of all groups where this customers belongs
		sqlCustGroups = " SELECT GROUP_ID,GROUP_NAME,GROUP_CID FROM Store_Customers_Groups WHERE "&_
						" GROUP_CID LIKE '%,"&CustId&",%' OR GROUP_CID LIKE '"&CustId&"' OR "&_
						" GROUP_CID LIKE '%,"&CustId&"' or GROUP_CID LIKE '"&CustId&",%' AND "&_
						" STORE_ID = "&Store_Id
		set rs = conn_store.execute(sqlCustGroups)
		
		if not rs.eof then
			while not rs.eof 
			CustIds = trim(rs("group_cid"))
			GroupId = 	trim(rs("group_id"))
			grpArray = split(CustIds,",")
			newGroupIds = ""
			
			for i = 0 to ubound(grpArray)
			' Remove the Id of the Current Customer
			if grpArray(i) <> CustId then 
				if trim(newGroupIds) <> "" then
				newGroupIds = newGroupIds&","&grpArray(i)
				else
				newGroupIds = newGroupIds&grpArray(i)
				end if
			end if
	
			next
			
			' update the new customer ids for groups
			sqlUpdate = " Update Store_Customers_Groups set group_cid = '"&newGroupIds&"' where "&_
						" Group_id = "&GroupId&" and Store_id = "&Store_Id
			conn_store.execute(sqlUpdate)
			rs.movenext
			wend
		end if
		rs.close
		set rs=nothing
		
		
		' Update with current selection
		newGroupIds = ""
		CustId = CustId
		newGroupIds = additional_groups
		
		if trim(newGroupIds) <> "" then
			sql = "select group_id,group_cid from Store_Customers_Groups where group_id in ("&newGroupIds&") and store_id = "&store_id
			session("sql")=sql
			set rs= conn_store.execute(sql)
			if not rs.eof then
				while not rs.eof
				custIds = trim(rs("group_cid"))
				grpId =  trim(rs("group_id"))
					if custIds <> "" then
						newGroupIds = custIds&","&custId
					else
						newGroupIds = custId
					end if
					sqlUpdate = " Update Store_Customers_Groups set group_cid = '"&newGroupIds&"' where "&_
								" Group_id = "&GrpId&" and Store_id = "&Store_Id	
				
					conn_store.execute(sqlUpdate)
				rs.movenext
				wend
			end if
		end if
	

		sql_update_customer = "Update store_customers Set Tax_Id='"&Tax_ID&"',Protected_page_access="&Protected_Access&", User_ID = '"&User_id&"' ,Password = '"&Password&"',Spam="&Spam&",Tax_Exempt = "&Tax_Exempt&" ,budget_orig = "&budget_orig&" , budget_left = "&budget_left&" ,reward_total = "&reward_total&", reward_left = "&reward_left&"  Where Cid = "&Cid&" and store_Id="&Store_Id
		session("sql") = sql_update_customer
		
		conn_store.Execute sql_update_customer
		if Record_type<>0 then
			Response.Redirect "Modify_customer.asp?cid="&cid&"&ssadr="&Record_type
		else
			Response.Redirect "Modify_customer.asp?cid="&cid
		end if
	end if

	'PROCESS SHIPPING CLASSES
	if Request.Form("Delete_Shipping_Method_Id") <> "" then
		'DELETE SHIPPING CLASS
		shipping_method_Id = Request.Form("Delete_Shipping_Method_Id")
		sql_delete_shipping_method="Delete from Store_Shippers_class_all where shipping_method_Id = "&shipping_method_Id&" and store_id="&Store_id
		conn_store.Execute sql_delete_shipping_method
		Response.Redirect Request.Form("Shipping_Form_Back_To")&sAddString
	end if 

	'ADD INSURANCE METHODS
	if request.form("Add_What_Insurance_Class") <> "" then
  		Insurance_Method_Name = checkStringForQ(Request.Form("Insurance_Method_Name"))
  		Insurance_class = Request.Form("Add_What_Insurance_Class")

  		select case Request.Form("Add_What_Insurance_Class")
  			case 1
  				Matrix_low = 0
  				Matrix_High = 0
  				Base_fee = cdbl(Request.Form("Base_Fee"))
  				Weight_fee = 0
  			case 4
  				Matrix_low = cdbl(Request.Form("Matrix_low"))
  				Matrix_High = cdbl(Request.Form("Matrix_High"))
  				Base_fee = cdbl(Request.Form("Base_Fee"))
  				Weight_fee = 0
  			case 5
  				Matrix_low = cdbl(Request.Form("Matrix_low"))
  				Matrix_High = cdbl(Request.Form("Matrix_High"))
  				Base_fee = cdbl(Request.Form("Base_Fee"))
  				Weight_fee = 0
  		end select
  		
  		if isNumeric(Matrix_low) and isNumeric(Matrix_High) and isNumeric(Weight_fee) and isNumeric(Base_fee) then
  				if Request.Form("op")="" then
  					sql_insert =  "insert into Store_Insurance_class_all (Store_id,Insurance_class,Insurance_method_name,Matrix_low,Matrix_High,Base_fee) Values ("&Store_id&","&Insurance_class&",'"&Insurance_Method_Name&"',"&Matrix_low&","&Matrix_High&","&Base_fee&")"
  					conn_store.Execute sql_insert
  					Response.Redirect Request.Form("Insurance_Form_Back_To")&sAddString
  				else
  					InsuranceClass=Request.Form("Add_What_Insurance_Class")
  					sqlUpdate="update Store_Insurance_class_all set Insurance_class=" & Insurance_class & ", " & _
  									"Insurance_Method_Name='" & Insurance_Method_Name & "', " & _
  									"Matrix_low=" & Matrix_low & ", " & _	
  									"Matrix_High=" & Matrix_High & ", " & _ 
  									"Base_fee=" & Base_fee & " " & _
  									" where " & _
  									"Insurance_class=" & InsuranceClass & " AND Store_id="&Store_id & " and " & _
  									"Insurance_method_Id=" & Request.Form("Insurance_Method_Id")
  					conn_store.Execute sqlUpdate
  					' back to where we came from ...
  					Response.Redirect Request.Form("Insurance_Form_Back_To")&sAddString
  				end if
  		else
  			Response.Redirect "admin_error.asp?message_id=1"
  		end if
	end if

	'ADD SHIPPING METHODS
	if request.form("Add_What_Shipping_Class") <> "" then
  		Shipping_Method_Name = checkStringForQ(Request.Form("Shipping_Method_Name"))
  		if instr(Shipping_Method_Name,"|") > 0 then
			fn_error "The name cannot contain a | symbol."
		end if
		Shippers_class = Request.Form("Add_What_Shipping_Class")
  		Countries = checkStringForQ(Request.Form("Countries"))
		
  		select case Request.Form("Add_What_Shipping_Class")
  			case 1
  				Matrix_low = 0
  				Matrix_High = 0
  				Base_fee = Request.Form("Base_Fee")
  				Weight_fee = 0
  			case 2
  				Matrix_low = Request.Form("Matrix_low")
  				Matrix_High = Request.Form("Matrix_High")
  				Base_fee = Request.Form("Base_Fee")
  				Weight_fee = Request.Form("Weight_fee")
  			case 3
  				Matrix_low = 0
  				Matrix_High = 0
  				Base_fee = Request.Form("Base_Fee")
  				Weight_fee = 0
  			case 4
  				Matrix_low = Request.Form("Matrix_low")
  				Matrix_High = Request.Form("Matrix_High")
  				Base_fee = Request.Form("Base_Fee")
  				Weight_fee = 0
  			case 5
  				Matrix_low = Request.Form("Matrix_low")
  				Matrix_High = Request.Form("Matrix_High")
  				Base_fee = Request.Form("Base_Fee")
  				Weight_fee = 0
  			case 6
  				Matrix_low = Request.Form("Matrix_low")
  				Matrix_High = Request.Form("Matrix_High")
  				Base_fee = Request.Form("Base_Fee")
  				Weight_fee = 0
  		end select
  		
  		Matrix_low = formatnumber(Matrix_Low,2,0,0,0)
		Matrix_High = formatnumber(Matrix_High,2,0,0,0)

		Base_fee = formatnumber(Base_Fee,2,0,0,0)
		Weight_fee = formatnumber(Weight_fee,2,0,0,0)
  		Zip_Start=request.form("Zip_Start")
  		Zip_End=request.form("Zip_End")
          Realtime_backup=request.form("Realtime_backup")
          if Realtime_backup="" then
          	Realtime_backup=0
          end if

  		Ship_Location_Id=request.form("Ship_Location_Id")

		' --------------------------------------------------------------------
		

  		if Is_In_Collection("1,2,3,4,5,6", Cstr(Request.Form("Add_What_Shipping_Class")),",") then
  			if isNumeric(Matrix_low) and isNumeric(Matrix_High) and isNumeric(Weight_fee) and isNumeric(Base_fee) then
  				if Request.Form("op")="" then
  					sql_insert =  "insert into Store_Shippers_class_all (Store_id,Shippers_class,Shipping_method_name,Matrix_low,Matrix_High,Base_fee,Weight_fee,Countries,Ship_Location_Id,Zip_Start,Zip_End,Realtime_backup) Values ("&Store_id&","&Shippers_class&",'"&Shipping_Method_Name&"',"&Matrix_low&","&Matrix_High&","&Base_fee&","&Weight_fee&",'"&Countries&"','"&Ship_Location_Id&"','"&Zip_Start&"','"&Zip_End&"',"&Realtime_backup&")"
  					
					conn_store.Execute sql_insert
  					Response.Redirect Request.Form("Shipping_Form_Back_To")&sAddString
  				else
  					ShipClass=Request.Form("Add_What_Shipping_Class")
  					sqlUpdate="update Store_Shippers_class_all set Shippers_class=" & Shippers_class & ", " & _
  									"Shipping_Method_Name='" & Shipping_Method_Name & "', " & _ 
  									"Realtime_backup=" & Realtime_backup & ", " & _
  									"Matrix_low=" & Matrix_low & ", " & _
  									"Matrix_High=" & Matrix_High & ", " & _
  									"Base_fee=" & Base_fee & ", " & _	
  									"Weight_fee=" & Weight_fee & ", " & _	
  									"Zip_Start='" & Zip_Start & "', " & _
  									"Zip_End='" & Zip_End & "', " & _
  									"Ship_Location_Id='" & Ship_Location_Id & "', " & _
  									"Countries='" & Countries & "' " & _
  									" where " & _
  									"Store_id="&Store_id & " and " & _
  									"shipping_method_Id=" & Request.Form("Shipping_Method_Id")

					conn_store.Execute sqlUpdate
  					' back to where we came from ...
  					Response.Redirect Request.Form("Shipping_Form_Back_To")&sAddString
  				end if
  			else
  				Response.Redirect "admin_error.asp?message_id=1"
  			end if
  		end if
	end if

	'DELETE COUSTOMER GROUP
	if Request.QueryString("Delete_Customer_Group_Id") <> "" then
		Group_Id = Request.QueryString("Delete_Customer_Group_Id")
		if isNumeric(Group_Id) then
			sql_delete="Delete from Store_Customers_Groups where Group_id = "&Group_id
			conn_store.Execute sql_delete
			Response.Redirect "customers_groups.asp"
		else
			 Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if

	'CREATE CUSTOMER GROUP
	if Request.Form("Create_Customer_Group") <> "" then 
		'RETRIEVE FORM DATA
		Group_Purchase_History = Request.Form("Group_Purchase_History")
		Group_budget_min = Request.Form("Group_budget_min")
		if isNumeric(Group_Purchase_History) and isNumeric(Group_budget_min) then
			Group_name = checkStringForQ(Request.Form("Group_name"))
			Group_Country = checkStringForQ(Request.Form("Group_Country"))
			Group_Dept = checkStringForQ(Request.Form("Group_Dept"))
			Group_City = checkStringForQ(Request.Form("Group_City"))
			Group_Company = checkStringForQ(Request.Form("Group_Company"))
			Group_Cid = checkStringForQ(Request.Form("Group_Cid"))
			Group_Cid_Array = Split(Group_Cid,",")
			for each cid in Group_Cid_Array
				if not isNumeric(cid) then
					fn_error "Customer ids must be numeric"
				end if
			next
			if Request.Form("op")<>"edit" then
				sql_insert = "insert into store_customers_groups (Store_id,Group_name,Group_Country,Group_City,Group_budget_min,Group_Company,Group_Purchase_History,Group_Cid) values ("&Store_id&",'"&Group_name&"','"&Group_Country&"','"&Group_City&"',"&Group_budget_min&",'"&Group_Company&"',"&Group_Purchase_History&",'"&Group_Cid&"') "
				conn_store.Execute sql_insert
				Response.Redirect "customers_groups.asp"
			else
				CustomerGroupId=Request.Form("Customer_Group_Id")
				sqlUpdate="update Store_Customers_Groups set Group_name='" & Group_name & "', " & _
						"Group_Country='" & Group_Country & "', " & _
						"Group_Dept='" & Group_Dept & "', " & _
						"Group_City='" & Group_City & "', " & _
						"Group_budget_min=" & Group_budget_min & ", " & _
						"Group_Company='" & Group_Company & "', " & _
						"Group_Purchase_History=" & Group_Purchase_History & ", " & _
						"Group_Cid='" & Group_Cid & "' " & _
						"where store_id=" & store_id & " and Group_Id=" & CustomerGroupId
				session("sql")=sqlUpdate
                                conn_store.Execute sqlUpdate
				Response.Redirect "customers_groups.asp"
			end if
		else
			Response.Redirect "admin_error.asp?message_id=1"
		end if
	end if

	if Request.Form("realtime") <> "" then
		Restrict_Options=checkstringforQ(request.form("Restrict_Options"))
		Countries = request.form("Countries")
		Max_Weight=request.form("Max_Weight")
		sql_update = "update store_real_time_settings set max_Weight="&Max_Weight&", countries='"&Countries&"',Restrict_Options='"&Restrict_Options&"' where store_id="&Store_Id
		session("sql") = sql_update
		conn_store.Execute sql_update
		Response.Redirect "realtime_options.asp"
	end if

 
	'DELETE A SHIPPING ADDRESS FROM ADDRESSBOOK
	if Request.Form("Delete_Ship_Addr") <> "" then
		cid = int(Request.QueryString("cid") )
		sql_delete = "delete from store_customers Where Cid = "&Cid&" And Record_type = "&Request.Form("Record_type")
		conn_store.Execute sql_delete
		if cstr(Request.Form("Record_type"))=cstr(1) then 
			sql_select_addrs = "SELECT max(Store_Customers.Record_type) as theMax FROM Store_Customers WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type <>  0"
			rs_Store.open sql_select_addrs,conn_store,1,1
			rs_Store.MoveFirst
			maxRT = rs_store("theMax")
			rs_store.close()
			sql_update = "update Store_Customers set Record_type=1 WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type ="&maxRT&" and Store_Id="&Store_Id
			conn_store.Execute sql_update
		end if
		Response.Redirect "Modify_customer.asp?cid="&cid&"&ssadr=1"
	end if

end if

%>
