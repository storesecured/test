<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="include/sub.asp"-->

<% 
on error resume next
'ERROR CHECKING

If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%
	
else
	returnTo = request.form("redirect")
	
	if returnTo = "" then
		returnTo = request.ServerVariables("HTTP_REFERER")
	end if
	select case Request.Form("Form_Name")

		'CAME FROM COMPANY WINDOW
		Case "company"
			Gift_Wrapping_Surcharge = Request.Form("Gift_Wrapping_Surcharge")
			if isNumeric(Gift_Wrapping_Surcharge) then
				Gift_Service = Request.Form("Gift_Service")
				if Gift_Service <> "" then
					Gift_Service = 1
				else
					Gift_Service = 0
				end if
				Gift_Message = Request.Form("Gift_Message")
				if Gift_Message <> "" then
					Gift_Message = 1
				else
					Gift_Message = 0
				end if
				
				Store_public = Request.Form("Store_public")
				if Store_public <> "" then
					Store_public = 1
				else
					Store_public = 0
				end if
				ExpressCheckout = Request.Form("ExpressCheckout")
				if ExpressCheckout <> "" then
					ExpressCheckout = 1
				else
					ExpressCheckout = 0
				end if
				Google_Checkout = Request.Form("Google_Checkout")
				Enable_IP_Tracking = Request.Form("Enable_IP_Tracking")
				if Enable_IP_Tracking <> "" then
					Enable_IP_Tracking = 1
				else
					Enable_IP_Tracking = 0
				end if
				No_Login = Request.Form("No_Login")
				if No_Login <> "" then
					No_Login = 1
				else
					No_Login = 0
				end if
				AllowCookies = Request.Form("AllowCookies")
				if AllowCookies <> "" then
					AllowCookies = 1
				else
					AllowCookies = 0
				end if	
				                 				
				When_Adding = Request.Form("When_Adding")
				if When_Adding = "" then
					When_Adding = 0
				else
					When_Adding = 1
				end if
				
				Cart_Thumbnails = Request.Form("Cart_Thumbnails")
				if Cart_Thumbnails = "" then
					Cart_Thumbnails = 0
				else
					Cart_Thumbnails = 1
				end if
				Save_Cart = Request.Form("Save_Cart")
				if Save_Cart = "" then
					Save_Cart = 0
				else
					Save_Cart = 1
				end if
				Hide_Retail_Price=request.form("Hide_Retail_Price")
				if Hide_Retail_Price = "" then
					Hide_Retail_Price=0
				else
					Hide_Retail_Price=1
				end if
				show_residential=request.form("show_residential")
				if show_residential = "" then
					show_residential=0
				else
					show_residential=1
				end if
				show_special_offers=request.form("show_special_offers")
				if show_special_offers = "" then
					show_special_offers=0
				else
					show_special_offers=1
				end if
				show_tax_exempt=request.form("show_tax_exempt")
				if show_tax_exempt = "" then
					show_tax_exempt=0
				else
					show_tax_exempt=1
				end if
				newsletter_receive=request.form("newsletter_receive")
				if newsletter_receive = "" then
					newsletter_receive=0
				else
					newsletter_receive=1
				end if
				Continue_Shopping = request.form("Continue_Shopping")
				Minimum_Amount = request.form("Minimum_Amount")
				if Minimum_Amount="" then
				Minimum_Amount=0.0
				end if
				
				Store_Zip= checkStringForQ(Request.Form("Store_Zip"))
				Store_State= checkStringForQ(Request.Form("Store_State"))
				Store_name= nullifyQ(Request.Form("Store_name"))
				Store_Company= checkStringForQ(Request.Form("Store_Company"))
				Store_Address1= checkStringForQ(Request.Form("Store_Address1"))
				Store_Address2= checkStringForQ(Request.Form("Store_Address2"))
				Store_City= checkStringForQ(Request.Form("Store_City"))
				Store_Phone= checkStringForQ(Request.Form("Store_Phone"))
				Store_Fax= checkStringForQ(Request.Form("Store_Fax"))
				Store_EMail= checkStringForQ(Request.Form("Store_Email"))
				Store_Country= checkStringForQ(Request.Form("Store_Country"))
				Store_Currency = checkStringForQ(Request.Form("Store_Currency"))
				
			    if (Store_Country="United States" or Store_Country="Canada") and (Store_State="") then
					    Response.Redirect "admin_error.asp?message_Add=<b>State</b> Field is Missing..."					
			    else
			        sql_Update_Store_Company =  "Update Store_Settings SET newsletter_receive="&newsletter_receive&",show_residential="&show_residential&",show_tax_exempt="&show_tax_exempt&",show_special_offers="&show_special_offers&",Enable_IP_Tracking="&Enable_IP_Tracking&",Cart_Thumbnails="&Cart_Thumbnails&",Continue_Shopping='"&Continue_Shopping&"',Save_Cart="&Save_Cart&", When_Adding="&When_Adding&", Gift_Message = "&Gift_Message&", Gift_Service = "&Gift_Service&", Gift_Wrapping_Surcharge = "&Gift_Wrapping_Surcharge&", Store_public="&Store_public&",ExpressCheckout="&ExpressCheckout&",No_Login="&No_Login&",AllowCookies="&AllowCookies&", Hide_Retail_Price=  " & Hide_Retail_Price&",Store_Currency = '"&Store_Currency&"', Store_name = '"&Store_name&"', Store_Company='"&Store_Company&"', Store_Address1='"&Store_Address1&"', Store_Address2='"&Store_Address2&"', Store_City='"&Store_City&"', Store_Zip='"&Store_Zip&"', Store_State='"&Store_State&"', Store_Country='"&Store_Country&"', Store_EMail='"&Store_EMail&"', Store_Phone='"&Store_Phone&"',Store_FAX='"&Store_FAX&"',Minimum_Amount=" & Minimum_Amount &" WHERE Store_id="&Store_id
			        conn_store.Execute sql_Update_Store_Company
			    end if
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
		'CAME FROM DOMAIN WINDOW
		Case "domain"
		
			sql_select = "select store_domain,store_password from store_settings WITH (NOLOCK) where store_id="&Store_Id
			rs_Store.open sql_select,conn_store,1,1
				Store_domain = rs_store("Store_Domain")
				Store_Password = rs_store("Store_Password")
			rs_Store.close
			
			if Store_Domain<>"" then
			   fn_error("You have already requested the domain name "&Store_Domain&".  You may only request one name.")
			end if

			if Request.Form("Transfer") <> "" then
				Subject = "Transfer"
				Store_Domain = Request.Form("Store_Domain1")
			else
				Subject = "Register"
				Store_Domain = Request.Form("Store_Domain")
				Store_Domain = Replace(Store_Domain,".com","")
        Store_Domain = Replace(Store_Domain,".net","")
        Store_Domain = Replace(Store_Domain,".info","")
        Store_Domain = Replace(Store_Domain,".org","")
        Store_Domain = Replace(Store_Domain,".biz","")
        Store_Domain = Replace(Store_Domain,".info","")
        Store_Domain = Replace(Store_Domain,".name","")
				Store_Domain = Store_Domain&Request.Form("domain_ext")
			end if
				
			StrDomain = lcase(Store_Domain)
			If Not StrDomain = "" Then
				DomainName = Trim(StrDomain)
				DomainName = Replace(DomainName,"http://","",vbTextCompare)
				DomainName = Replace(DomainName,"www.","",vbTextCompare)

				Set whoisdll = Server.CreateObject("WhoisDLL.Whois")
				whoisdll.WhoisServer = "whois.internic.net"
				Result = ""
				Result = whoisdll.whois(Trim(DomainName))

				StartPos = InStr(Result,"Whois Server:")
				If StartPos > 0 Then
					StartPos = StartPos + Len("Whois Server:")
					EndPos = InStr(StartPos,Result,vblf)
					WhoisServer = Trim(Mid(Result,StartPos,EndPos-StartPos))
					Set whoisdll = Server.CreateObject("WhoisDLL.Whois") 
					whoisdll.WhoisServer = WhoisServer
					Result = ""
					Result = whoisdll.whois(Trim(DomainName))
					WhoisResult = Result
					iDomainAvailable = 0
				Else
					If InStr(1,Result,"No match for",vbTextCompare) > 0 Then
					 iDomainAvailable = 1
				  end if
				End If
			End If

			if Subject = "Register" and iDomainAvailable = 0 then
			   response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("This domain name cannot be registered because it is already registered.<BR><BR>If you own this domain please request a transfer instead of registration.<BR><BR>If you do not own this domain please choose another name that is not already taken.")
			elseif Subject = "Transfer" then 'and iDomainAvailable = 1 then

				'***********************
								Darray =  split(store_domain,".")
								qualifier= Darray(ubound(Darray))
								if len(qualifier) >1 and len(qualifier) <5 then

								else
									response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("The domain name does not have a valid extension, ie .com, .net, .org etc.")
								end if
				'***********************



		end if

			sql_select = "select store_id from store_settings WITH (NOLOCK) where site_name like'%"&Store_Domain&"%'  or store_domain like '%"&Store_Domain&"%' or store_domain2 like '%"&Store_Domain&"%'"
			rs_Store.open sql_select,conn_store,1,1
			if not rs_Store.bof and not rs_Store.eof then
			   rs_Store.close
			   response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("This domain name is already in use.")
			end if
			rs_Store.close
			
			host = replace(lcase(request.servervariables("HTTP_HOST")),"manage.","")

			sIPAddress =Server_IP

			sql_Update_Store_Company =  "Update Store_Settings SET Store_Domain='"&Store_Domain&"',mail_done=0 WHERE Store_id="&Store_id
			conn_store.Execute sql_Update_Store_Company

      SupportRequest=Subject & " domain"
			if Subject <> "" then
				Body = Subject & " domain: " & Store_Domain & " for store " & Store_id
				Subject1 = Subject & " " & Store_Domain & " " & Store_Id
				if Subject = "Register" then
				   sql_select = "select * from sys_billing where store_id="&Store_Id
				   rs_store.open sql_select,conn_store,1,1
				   if not rs_Store.bof and not rs_Store.eof then
                                                sRegister = rs_Store("first_name") & " " & rs_Store("last_name") & vbcrlf &_
                                                rs_Store("Address") & vbcrlf &_
                                                rs_Store("City") & ", " & rs_Store("State") & " " & rs_Store("Zip") &_
                                                rs_Store("Country") & vbcrlf &_
                                                rs_Store("Phone") & vbcrlf &_
                                                rs_Store("Fax") & vbcrlf &_
                                                rs_Store("Email")
                                   else
                                       sRegister = "Register under default name, no addtl info."
                                   end if
				   rs_store.close
			           
                                end if
                        end if

			if Store_Domain <> "" then
				Store_Domain = replace(Store_Domain,"www.","")
				sql_insert = "New_Support_Request "&Store_Id&",'"&SupportRequest&"','"&Store_Domain&" IP Address: "&sIPAddress&"','Customer','"&Store_Email&"'"
				conn_store.Execute sql_insert

			end if

			response.redirect "transfer.asp"



		'***********ADDING ADDITIONAL DOMAIN TO THE STORE*****************
		case "Additional_Domain"
			sql_select = "select store_domain2,store_password from store_settings WITH (NOLOCK) where store_id="&Store_Id
			rs_Store.open sql_select,conn_store,1,1
			Store_domain2 = rs_store("Store_Domain2")
			Store_Password = rs_store("Store_Password")
			rs_Store.close

			if store_domain2<>"" then
			   fn_error("You have already requested the domain name "&store_domain2&".  You may only request one name.")
			end if

             Subject = "Transfer"
                        store_domain2 = request.Form("Store_Addl_Domain")
			'***********CHECKING VALIDITY

								Darray =  split(store_domain2,".")
								qualifier= Darray(ubound(Darray))
								if len(qualifier) >1 and len(qualifier) <5 then

								else
									response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("The domain name does not have a valid extension, ie .com, .net, .org etc.")
								end if

								sql_select = "select store_id from store_settings WITH (NOLOCK) where site_name like '%"&store_domain2&"%' or  store_domain like '%"&store_domain2&"%' or store_domain2 like '%"&store_domain2&"%'"

								rs_Store.open sql_select,conn_store,1,1
								if not rs_Store.bof and not rs_Store.eof then
								   rs_Store.close
								   response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("This domain name is already in use.")
								end if
								rs_Store.close


'			response.end
			'**********************
                        StrDomain = lcase(store_domain2)
				         sIPAddress =Server_IP

			SupportRequest=Subject & " domain"
			Body = Subject & " domain: " & store_domain2 & " for store " & Store_id
			Subject1 = Subject & " " & store_domain2 & " " & Store_Id

			addl_domain = store_domain2
			update_sql = "update store_settings set store_domain2='"  & addl_domain & "',mail_done=0 where store_id= "&Store_Id
			conn_store.execute update_sql

			store_domain2 = replace(addl_domain,"www.","")
			sql_insert = "New_Support_Request "&Store_Id&",'"&SupportRequest&"','"&store_domain2&" IP Address: "&sIPAddress&"','Customer','"&Store_Email&"'"
			conn_store.Execute sql_insert

		'***************END OF ADDITIONAL DOMAIN**************************


		'CAME FROM INVENTORY DISPLAY WINDOW
		Case "Inventory_display"
			dept_display = Request.Form("dept_display")
			item_display = Request.Form("item_display")
			item_f_display = Request.Form("item_f_display")
			item_rows = Request.Form("item_rows")
			item_f_rows = Request.Form("item_f_rows")
			dept_rows = Request.Form("dept_rows")
			if isNumeric(dept_display) and isNumeric(item_display) and isNumeric(item_rows) and isNumeric(item_f_rows) and isNumeric(dept_rows) and isNumeric(item_f_display) then
				if (item_rows * (item_display+1)) > 50 then
				  fn_error("You may not specify more than 50 items on a page.")
				end if
				sql_Update_inventory_display = "Update Store_Settings set item_f_display="&item_f_display&", dept_display = "&dept_display&", item_display="&item_display&", item_f_rows="&item_f_rows&", item_rows="&item_rows&", dept_rows="&dept_rows&" where Store_id = "&Store_id
				conn_store.Execute sql_Update_inventory_display 
				
				server.execute "reset_design.asp"
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if

		'CAME FROM DESGINER TEMPLATE WINDOW
		Case "Designer_template"
			Store_design = Request.Form("Store_header")
			Store_design=replace(Store_design,Site_Name2&"images","/images")
         Store_design=replace(replace(Store_design,"<OBJ_TEXTAREA_START","<TEXTAREA"),"<OBJ_TEXTAREA_END","</TEXTAREA")
         Store_design=replace(Store_design,"</%OBJ_TOP_DES_OBJECTS_OBJ%"&">","")
         Store_design=replace(Store_design,"</%OBJ_LEFT_DES_OBJECTS_OBJ%"&">","")
			Store_design=replace(Store_design,"</%OBJ_RIGHT_DES_OBJECTS_OBJ%"&">","")
			Store_design=replace(Store_design,"</%OBJ_CENTOP_DES_OBJECTS_OBJ%"&">","")
			Store_design=replace(Store_design,"</%OBJ_CENBOT_DES_OBJECTS_OBJ%"&">","")
			Store_design=replace(Store_design,"</%OBJ_BOTTOM_DES_OBJECTS_OBJ%"&">","")
         Store_design=replace(Store_design,"</%OBJ_NAV_BUTTONS_OBJ%"&">","")
         Store_design=replace(Store_design,"</%OBJ_NAV_LINKS_OBJ%"&">","")
         Store_design=replace(Store_design,Site_Name,"OBJ_SWITCH_NAME")
         Store_design=replace(Store_design,"<TBODY></TBODY>","")
         Store_design=replace(Store_design,"bgcolor=""#000bc0"" ","bgcolor=""%OBJ_BG_COLOR_OBJ%"" ")

			
			cPos = InStr(Store_design, "%OBJ_CENTER_CONTENT_OBJ%")
			if cPos > 0 then
				Design_Template_Id=Request.Form("Template_Id")
                template_html = nullifyQ(Store_design)
                template_head = Request.Form("template_head")

			    template_head = nullifyQ(template_head)
			
				sql_update_template = "Update store_design_template set template_html = '"&template_html&"',template_head='"&template_head&"' where Store_id = "&Store_id& " and Template_Id="&Design_Template_Id
				conn_store.execute sql_update_template

				response.redirect "designer_template.asp?Id="&Design_Template_Id
				
			else
				Response.Redirect "admin_error.asp?message_id=102"
			end if
			
		'CUSTOMIZE INVOICE HEADER
		Case "Customize_invoice"
		
			Invoice_header = image_replace(Request.Form("Invoice_header"))
			Invoice_footer = image_replace(Request.Form("Invoice_footer"))

			
			Invoice_Header = replace(Invoice_Header,Site_Name2&"images","/images")
            Invoice_Header=replace(replace(Invoice_Header,"<OBJ_TEXTAREA_START","<TEXTAREA"),"<OBJ_TEXTAREA_END","</TEXTAREA")
            
            Invoice_Footer = replace(Invoice_Footer,Site_Name2&"images","/images")
            Invoice_Footer=replace(replace(Invoice_Footer,"<OBJ_TEXTAREA_START","<TEXTAREA"),"<OBJ_TEXTAREA_END","</TEXTAREA")
    
			Invoice_header = Replace(Invoice_Header,"'","''")
			Invoice_footer = Replace(Invoice_Footer,"'","''")
				
				sql_Update_Store_Activation = "Update Store_settings set Invoice_header = '"&Invoice_header&"', Invoice_footer = '"&Invoice_footer&"' where Store_id = "&Store_id
				conn_store.Execute sql_Update_Store_Activation
			
			
	
		Case "activation"
			Store_Active = Request.Form("Store_Active")
			if isNumeric(Store_Active) then
				Store_Close_Message = nullifyQ(Request.Form("Store_Close_Message"))
				sql_update_instance = "exec wsp_page_content_top "&Store_id&",4,'"&Store_Close_Message&"';"
			        conn_store.Execute sql_update_instance
                                sql_Update_Store_Activation = "Update Store_Settings set Store_active = "&Store_active&" where Store_id = "&Store_id
				conn_store.Execute sql_Update_Store_Activation
				
				
				
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if

		'CAME FROM item SETTINGS WINDOW
		Case "item_settings"
				dept_display = Request.Form("dept_display")
				item_display = Request.Form("item_display")
				item_f_display = Request.Form("item_f_display")
				item_rows = Request.Form("item_rows")
				item_f_rows = Request.Form("item_f_rows")
				dept_rows = Request.Form("dept_rows")
				if isNumeric(dept_display) and isNumeric(item_display) and isNumeric(item_rows) and isNumeric(item_f_rows) and isNumeric(dept_rows) and isNumeric(item_f_display) then
					if (item_rows * (item_display+1)) > 50 then
					  response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("You may not specify more than 50 items on a page.")
					end if
		
					Inventory_Reduce = Request.Form("Inventory_Reduce")
					if Inventory_Reduce = "" then
						Inventory_Reduce = 0
					else
						Inventory_Reduce = 1
					end if
					
					Unverified_Reduce = Request.Form("Unverified_Reduce")
					if Unverified_Reduce = "" then
						Unverified_Reduce = 0
					else
						Unverified_Reduce = 1
					end if
					
					Show_Jump = Request.Form("Show_Jump")
					if Show_Jump = "" then
						Show_Jump = 0
					else
						Show_Jump = 1
					end if
					
					Hide_Empty_Depts = Request.Form("Hide_Empty_Depts")
					if Hide_Empty_Depts = "" then
						Hide_Empty_Depts = 0
					else
						Hide_Empty_Depts = 1
					end if
					
					Hide_OutofStock_Items = Request.Form("Hide_OutofStock_Items")
					if Hide_OutofStock_Items = "" then
						Hide_OutofStock_Items = 0
					else
						Hide_OutofStock_Items = 1
					end if
									
					Show_TopNav = Request.Form("Show_TopNav")
					if Show_TopNav = "" then
						Show_TopNav = 0
					else
						Show_TopNav = 1
					end if

					'added on 05-17-2005
					'For Item-Navigation
					
					Show_ItemNav = Request.Form("Show_ItemNav")
					if Show_ItemNav = "" then
						Show_ItemNav = 0
					else
						Show_ItemNav = 1
					end if
					
					Show_SubDept = Request.Form("Show_SubDept")
					if Show_SubDept = "" then
						Show_SubDept = 0
					else
						Show_SubDept = 1
					end if
					
					Reload_Attr = Request.Form("Reload_Attr")
					if Reload_Attr = "" then
						Reload_Attr = 0
					else
						Reload_Attr = 1
					end if
							
					User_Defined_Fields = request.form("User_Defined_Fields")
					User_Defined_Fields_2 = request.form("User_Defined_Fields_2")
					User_Defined_Fields_3 = request.form("User_Defined_Fields_3")
					User_Defined_Fields_4 = request.form("User_Defined_Fields_4")
					User_Defined_Fields_5 = request.form("User_Defined_Fields_5")

					sql_Update_Store_Activation = "Update Store_Settings set reload_attr="&reload_attr&", item_f_display="&item_f_display&", dept_display = "&dept_display&", item_display="&item_display&", item_f_rows="&item_f_rows&", item_rows="&item_rows&", dept_rows="&dept_rows&", Hide_OutofStock_Items="&Hide_OutofStock_Items&",User_Defined_Fields='"&User_Defined_Fields&"',User_Defined_Fields_2='"&User_Defined_Fields_2&"',User_Defined_Fields_3='"&User_Defined_Fields_3&"',User_Defined_Fields_4='"&User_Defined_Fields_4&"',User_Defined_Fields_5='"&User_Defined_Fields_5&"',Unverified_Reduce="&Unverified_Reduce&",Show_TopNav="&Show_TopNav&",Detail_NextPrev="&Show_ItemNav&", Show_Jump="&Show_Jump&", Hide_Empty_Depts="&Hide_Empty_Depts&", Inventory_Reduce = "&Inventory_Reduce&", subdept_location="&Show_SubDept&"  where Store_id = "&Store_id
					conn_store.Execute sql_Update_Store_Activation
                                
					server.execute "reset_design.asp"
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
		'CAME FROM AFFILIATES WINDOW
		Case "Affiliate_settings"
			payPercent = Request.Form("payPercent")
			payAmount = Request.Form("payAmount")
			Affiliate_payout = Request.Form("Affiliate_payout")
			Affiliate_cookie = Request.Form("Affiliate_cookie")
			if isNumeric(Affiliate_payout) and isNumeric(Affiliate_cookie) then
				Enable_affiliates = Request.Form("Enable_affiliates")
				if Enable_affiliates <> "" then
					Enable_affiliates = 1
				else
					Enable_affiliates = 0
				end if
	
				Screen_affiliates = Request.Form("Screen_affiliates")
				if Screen_affiliates <> "" then
					Screen_affiliates = 1
				else
					Screen_affiliates = 0
				end if

				Affiliate_type = Request.Form("Affiliate_type")
				if Affiliate_type = 1 then
					Affiliate_amount = payPercent
				else
					Affiliate_amount = payAmount
				end if
	
				sql_Update_Store_Affiliate = "Update Store_Settings set Affiliate_Cookie="&Affiliate_cookie&", Enable_affiliates = "&Enable_affiliates&", Screen_affiliates = "&Screen_affiliates&", Affiliate_type="&Affiliate_type&",Affiliate_payout="&Affiliate_payout&",Affiliate_amount="&Affiliate_amount&" where Store_id = "&Store_id
				conn_store.Execute sql_Update_Store_Affiliate
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
		'CAME FROM AFFILIATES WINDOW
		
		Case "coupon_settings"
			Show_Coupon = Request.Form("Show_Coupon")
			if Show_Coupon = "" then
				Show_Coupon = 0
			else
				Show_Coupon = 1
			end if    

			sql_Update = "Update Store_Settings set Show_Coupon="&Show_Coupon&" where Store_id = "&Store_id
			conn_store.Execute sql_Update
			
		Case "reward_settings"
			Rewards_Minimum = Request.Form("Rewards_Minimum")
			Rewards_Percent = Request.Form("Rewards_Percent")
			if isNumeric(Rewards_Minimum) and isNumeric(Rewards_Percent) then
				Enable_Rewards = Request.Form("Enable_Rewards")
				if Enable_Rewards <> "" then
					Enable_Rewards = 1
				else
					Enable_Rewards = 0
				end if

				sql_Update_Store_Rewards = "Update Store_Settings set Enable_Rewards = "&Enable_Rewards&", Rewards_Percent = "&Rewards_Percent&", Rewards_Minimum="&Rewards_Minimum&" where Store_id = "&Store_id
				conn_store.Execute sql_Update_Store_Rewards
				
			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if
		'CAME FROM SHIPPING CLASS WINDOW
		Case "Shipping_Class"
			Shipping_Class = Request.Form("Shipping_Class")&","&Request.Form("Shipping_Class_Real")
			RTO = ""
			if request.form("USE_UPS")<>"" then
				RTO = RTO&", USE_UPS=-1"
				USE_UPS=-1
			else
				RTO = RTO&", USE_UPS=0"
				USE_UPS=0
			end if
			if request.form("USE_USPS")<>"" then
				RTO = RTO&", USE_USPS=-1"
				USE_USPS=-1
			else
				RTO = RTO&", USE_USPS=0"
				USE_USPS=0
			end if
			if request.form("USE_DHL")<>"" then
				RTO = RTO&", USE_DHL=-1"
				USE_DHL=-1
				DHL_Service=request.form("DHL_Service")
			else
				RTO = RTO&", USE_DHL=0"
				USE_DHL=0
			end if
			if request.form("USE_FEDEX")<>"" then
				RTO = RTO&", USE_FEDEX=-1"
				USE_FEDEX=-1
			else
				RTO = RTO&", USE_FEDEX=0"
				USE_FEDEX=0
			end if

			if request.form("USE_CANADA")<>"" then
				RTO = RTO&", USE_CANADA=-1"
				USE_CANADA=-1
			else
				RTO = RTO&", USE_CANADA=0"
				USE_CANADA=0
			end if

			if request.form("USE_AIRBORNE")<>"" then
				RTO = RTO&", USE_AIRBORNE=-1"
				USE_AIRBORNE=-1
			else
				RTO = RTO&", USE_AIRBORNE=0"
				USE_AIRBORNE=0
			end if
			
               if request.form("USE_UPS")<>"" then
					if request.form("UPS_User")="" or request.form("UPS_Password")="" or request.form("UPS_AccessLicense")="" then
						Response.Redirect "admin_error.asp?message_id=85"
					end if
				end if
				
				if request.form("Fedex_Ground") <>"" then
				   Fedex_Ground="True"
				else
				   Fedex_Ground="False"
				end if
				Restrict_Options=checkstringforQ(request.form("Restrict_Options"))
				Countries = request.form("Countries")
				Max_Weight=request.form("Max_Weight")
				Show_Shipping = Request.Form("Show_Shipping")
				if Show_Shipping = "" then
					Show_Shipping = 0
				else
					Show_Shipping = 1
				end if
                                Ship_Multi = Request.Form("Ship_Multi")
				if Ship_Multi <> "" then
					Ship_Multi = 1
				else
					Ship_Multi = 0
				end if
				sHandling_Fee = request.form("Handling_Fee")
			        sHandling_Weight = request.form("Handling_Weight")


				sql_Update_Store_Real_Time = "Update Store_Real_Time_Settings set Fedex_DropoffType ='"&request.form("FedEx_DropType")&"',max_Weight="&Max_Weight&", countries='"&Countries&"',Restrict_Options='"&Restrict_Options&"',Canada_Login='"&request.form("Canada_Login")&"', UPS_Pickup='"&request.form("UPS_Pickup")&"', UPS_Pack='"&request.form("UPS_Pack")&"', Fedex_Pack='"&request.form("Fedex_Pack")&"', Fedex_Ground='"&Fedex_Ground&"', UPS_AccessLicense='"&request.form("UPS_AccessLicense")&"', UPS_User='"&request.form("UPS_User")&"', UPS_Password='"&request.form("UPS_Password")&"', UPS_Account_Number='"&request.form("UPS_Account_Number")&"', UPS_Shipper_Name='"&request.form("UPS_Shipper_Name")&"',DHL_Service='"&DHL_Service&"', Fedex_Account_Number='"&request.form("Fedex_Account_Number")&"',Fedex_Meter_Number='"&request.form("Fedex_Meter_Number")&"',Fedex_Transaction_Identifier='"&request.form("Fedex_Transaction_Identifier")&"',Fedex_Carrier_Code='"&request.form("Fedex_Carrier_Code")&"',Fedex_Services='"&request.form("Fedex_Services")&"' where Store_id = "&Store_id
                                conn_store.Execute sql_Update_Store_Real_Time
                                sql_Update_Store_Real_Time = "Update Store_Settings set shipping_classes='"&shipping_class&"',handling_fee="&shandling_fee&", handling_weight="&sHandling_Weight&", ship_multi="&ship_multi&", show_shipping="&show_shipping&" where Store_id = "&Store_id
                                conn_store.Execute sql_Update_Store_Real_Time
                                


			
			' --------------------------
			sql_Update_Store_Activation = "Update Store_Settings set Shipping_Classes = '"&Shipping_Class&"' "&RTO&" where Store_id = "&Store_id
			conn_store.Execute sql_Update_Store_Activation

		'CAME FROM PAYMENT METHODS WINDOW
		Case "Store_Payment"
			store_payment_id=request.form("store_payment_id")
			payment_name=checkstringforQ(request.form("payment_name"))
			Accept=request.form("Accept")
			payment_method_message=checkstringforQ(request.form("payment_method_message"))
			if Accept <> "" then
				Accept = 1
			else
				Accept = 0
			end if

			if store_payment_id="" then
				sql_update_payment_methods = "insert into store_payment_methods (accept,payment_name,payment_method_message,store_id) values ("&accept&",'"&payment_name&"','"&payment_method_message&"',"&store_id&")"
			else
				sql_update_payment_methods = "UPDATE Store_Payment_methods SET Accept = "&accept&",payment_name='"&payment_name&"',payment_method_message='"&payment_method_message&"' WHERE store_payment_id ="&store_payment_id&" AND Store_id="&Store_id&";"
			end if 

			conn_store.Execute sql_update_payment_methods
			response.redirect "payment_manager.asp"
			

		Case "Real_Time_Settings"
			Real_Time_Processor = Request.Form("Real_Time_Processor")
			Auth_Capture = Request.Form("Auth_Capture")
			Use_CVV2 = Request.Form("Use_CVV2")
			Paypal_Express = Request.Form("Paypal_Express")
			Google_Checkout = Request.Form("Google_Checkout")
		
                	if Request.Form("GoogleCheckout_ButtonStyle")<>"" then
			GoogleCheckout_ButtonStyle= Request.Form("GoogleCheckout_ButtonStyle")
			else
			GoogleCheckout_ButtonStyle=0
			end if

			if Auth_Capture = "" then
				Auth_Capture = 0
			else
				Auth_Capture = 1
			end if
			if Paypal_Express = "" then
				Paypal_Express = 0
			else
				Paypal_Express = 1
			end if
			if Google_Checkout = "" then
				Google_Checkout = 0
			else
				Google_Checkout = 1
			end if
			if Use_CVV2 = "" then
				Use_CVV2 = 0
			else
				Use_CVV2 = 1
			end if
			Show_SecureLogo = Request.Form("Show_SecureLogo")
			if Show_SecureLogo = "" then
				Show_SecureLogo = 0
			else
				Show_SecureLogo = 1
			end if	


			if isNumeric(Real_Time_Processor) then

                Pay_To = Encrypt(checkStringForQ(request.form("Pay_To")))
                Pay_Currency = Encrypt(checkStringForQ(request.form("Pay_Currency")))
                sql_update="exec wsp_real_time_update "&store_id&",4,'Pay_To','"&Pay_To&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",4,'Pay_Currency','"&Pay_Currency&"';"

                PaytoID = Encrypt(checkStringForQ(request.form("PaytoID")))
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",33,'PaytoID','"&PaytoID&"';"

                PayPal_Pro_Api_Username =  Encrypt(checkStringForQ(request.form("PayPal_Pro_Api_Username")))
                PayPal_Pro_Password =  Encrypt(checkStringForQ(request.form("PayPal_Pro_Password")))
                PayPal_Pro_Currency =  Encrypt(checkStringForQ(request.form("PayPal_Pro_Currency")))
                PayPal_Pro_Signature =  Encrypt(checkStringForQ(request.form("PayPal_Pro_Signature")))
                Certificate_file_name =  Encrypt(checkStringForQ(request.form("Certificate_file_name")))
                	 
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",36,'PayPal_Pro_Api_Username','"&PayPal_Pro_Api_Username&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",36,'PayPal_Pro_Password','"&PayPal_Pro_Password&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",36,'PayPal_Pro_Currency','"&PayPal_Pro_Currency&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",36,'PayPal_Pro_Signature','"&PayPal_Pro_Signature&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",36,'Certificate_file_name','"&Certificate_file_name&"';"

                '****Google checkout*******
                Merchant_ID = Encrypt(checkStringForQ(request.form("Merchant_ID")))
                Merchant_Key = Encrypt(checkStringForQ(request.form("Merchant_Key")))
                Merchant_Currency = Encrypt(checkStringForQ(request.form("Merchant_Currency")))
                
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",38,'Merchant_ID','"&Merchant_ID&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",38,'Merchant_Key','"&Merchant_Key&"';"
                sql_update=sql_update&"exec wsp_real_time_update "&store_id&",38,'Merchant_Currency','"&Merchant_Currency&"';"

                if Real_Time_Processor = 1 then
                    publisher_name = Encrypt(checkStringForQ(Request.Form("publisher-name")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'publisher-name','"&publisher_name&"';"

                elseif Real_Time_Processor = 2 then
                    x_Login = Encrypt(checkStringForQ(Request.Form("x_Login")))
                    x_tran_key = Encrypt(checkStringForQ(Request.Form("x_tran_key")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'x_Login','"&x_Login&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'x_tran_key','"&x_tran_key&"';"
                elseif Real_Time_Processor = 5 then
                    PsiGate_StoreID= Encrypt(checkStringForQ(request.form("PsiGate_StoreID")))
                    Passphrase= Encrypt(checkStringForQ(request.form("Passphrase")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'PsiGate_StoreID','"&PsiGate_StoreID&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Passphrase','"&Passphrase&"';"
                elseif Real_Time_Processor = 7 then
                    merchant_echo_id = Encrypt(checkStringForQ(request.form("merchant_echo_id")))
                    merchant_echo_pin = Encrypt(checkStringForQ(request.form("merchant_echo_pin")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'merchant_echo_id','"&merchant_echo_id&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'merchant_echo_pin','"&merchant_echo_pin&"';"
                elseif Real_Time_Processor = 8 then
                    pl_Login = Encrypt(checkStringForQ(request.form("pl_Login")))
                    pl_Partner = Encrypt(checkStringForQ(request.form("pl_Partner")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'pl_Login','"&pl_Login&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'pl_Partner','"&pl_Partner&"';"    
                elseif Real_Time_Processor = 9 then
                    v_User = Encrypt(checkStringForQ(request.form("v_User")))
                    v_Vendor = Encrypt(checkStringForQ(request.form("v_Vendor")))
                    v_Partner = Encrypt(checkStringForQ(request.form("v_Partner")))
                    v_Password = Encrypt(checkStringForQ(request.form("v_Password")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'v_User','"&v_User&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'v_Vendor','"&v_Vendor&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'v_Partner','"&v_Partner&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'v_Password','"&v_Password&"';"
                elseif Real_Time_Processor = 10 then
                    storename = Encrypt(checkStringForQ(request.form("storename")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'storename','"&storename&"';"
                elseif Real_Time_Processor = 11 then
                    seller_id = Encrypt(checkStringForQ(request.form("seller_id")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'seller_id','"&seller_id&"';"
                elseif Real_Time_Processor = 12 then
                    MerchantNumber = Encrypt(checkStringForQ(request.form("MerchantNumber")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'MerchantNumber','"&MerchantNumber&"';"
                elseif Real_Time_Processor = 14 then
                    bmc_login = Encrypt(checkStringForQ(request.form("bmc_login")))
                    account_id = Encrypt(checkStringForQ(request.form("account_id")))
                    secret_key = Encrypt(checkStringForQ(request.form("secret_key")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'bmc_login','"&bmc_login&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'account_id','"&account_id&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'secret_key','"&secret_key&"';"
                elseif Real_Time_Processor = 15 then
                    ePNAccount = Encrypt(checkStringForQ(request.form("ePNAccount")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'ePNAccount','"&ePNAccount&"';"
                elseif Real_Time_Processor = 17 then
                    WorldPay_InstId = Encrypt(checkStringForQ(request.form("WorldPay_InstId")))
                    WorldPay_Currency = Encrypt(checkStringForQ(request.form("WorldPay_Currency")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'WorldPay_InstId','"&WorldPay_InstId&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'WorldPay_Currency','"&WorldPay_Currency&"';"
                elseif Real_Time_Processor = 19 then
                    VendorName = Encrypt(checkStringForQ(request.form("VendorName")))
                    VendorPassword = Encrypt(checkStringForQ(request.form("VendorPassword")))
                    VendorCurrency = Encrypt(checkStringForQ(request.form("VendorCurrency")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'VendorName','"&VendorName&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'VendorPassword','"&VendorPassword&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'VendorCurrency','"&VendorCurrency&"';"
                elseif Real_Time_Processor = 20 then
                    mb_pay_to_email = Encrypt(checkStringForQ(request.form("mb_pay_to_email")))
                    mb_pay_currency	= Encrypt(checkStringForQ(request.form("mb_pay_currency")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'mb_pay_to_email','"&mb_pay_to_email&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'mb_pay_currency','"&mb_pay_currency&"';"
                elseif Real_Time_Processor = 22 then
                    nochex_email = Encrypt(checkStringForQ(request.form("nochex_email")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'nochex_email','"&nochex_email&"';" 
                elseif Real_Time_Processor = 24 then
                    MerchantID = Encrypt(checkStringForQ(request.form("MerchantID")))
                    MerchantKey = Encrypt(checkStringForQ(request.form("MerchantKey")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'MerchantID','"&MerchantID&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'MerchantKey','"&MerchantKey&"';"
                elseif Real_Time_Processor = 26 then
                    M_ID = Encrypt(checkStringForQ(request.form("M_ID")))
                    M_Key = Encrypt(checkStringForQ(request.form("M_Key")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'M_ID','"&M_ID&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'M_Key','"&M_Key&"';"
                elseif Real_Time_Processor = 28 then
                    Cybersource_Id = Encrypt(checkStringForQ(request.form("Cybersource_Id")))
                    Cybersource_Currency = Encrypt(checkStringForQ(request.form("Cybersource_Currency")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Cybersource_Id','"&Cybersource_Id&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Cybersource_Currency','"&Cybersource_Currency&"';"
                elseif Real_Time_Processor = 29 then
                    Xor_User = Encrypt(checkStringForQ(request.form("Xor_User")))
                    Xor_Password = Encrypt(checkStringForQ(request.form("Xor_Password")))
                    Xor_Currency = Encrypt(checkStringForQ(request.form("Xor_Currency")))
                    Xor_Terminalno = Encrypt(checkStringForQ(request.form("Xor_Terminalno")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Xor_User','"&Xor_User&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Xor_Password','"&Xor_Password&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Xor_Currency','"&Xor_Currency&"';"
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Xor_Terminalno','"&Xor_Terminalno&"';"
                elseif Real_Time_Processor = 31 then
                    Propay_username = Encrypt(checkStringForQ(request.form("Propay_username")))
                    sql_update=sql_update&"exec wsp_real_time_update "&store_id&","&Real_Time_Processor&",'Propay_username','"&Propay_username&"';"



                end if

				sql_update = sql_update&"Update Store_Settings set Paypal_Express="&Paypal_Express&",GoogleCheckout=" & Google_Checkout & ",GoogleCheckout_ButtonStyle=" & GoogleCheckout_ButtonStyle & ",Show_SecureLogo="&Show_SecureLogo&", Real_Time_Processor = "&Real_Time_Processor&",Use_CVV2="&Use_CVV2&", Auth_Capture="&Auth_Capture&" where Store_id = "&Store_id
				conn_store.Execute sql_update


			else
				Response.Redirect "admin_error.asp?message_id=1"
			end if

		'CAME FROM STORE_EMAILS WINDOW
		case "Emails" 
			Html_Notifications = Request.Form("Html_Notifications")
			if Html_Notifications = "" then
				Html_Notifications = 0
			else
				Html_Notifications = 1
			end if	

			if Request.Form("Registration_enable") <> "" then
				Registration_enable = 1
			else
				Registration_enable = 0
			end if
			if Request.Form("Registration_user_psw") <> "" then
				Registration_user_psw = 1
			else
				Registration_user_psw = 0
			end if	
			if Request.Form("order_notification_invoice") <> "" then
				order_notification_invoice = 1
			else
				order_notification_invoice = 0
			end if	
			if Request.Form("order_notification_to_admin_enable") <> "" then
				order_notification_to_admin_enable = 1
			else
				order_notification_to_admin_enable = 0
			end if
			if Request.Form("order_notification_enable") <> "" then
				order_notification_enable = 1
			else
				order_notification_enable = 0
			end if
			Registration_subject = replace(Request.Form("Registration_subject"),"'","''")
			Registration_body = replace(Request.Form("Registration_body"),"'","''")
			Registration_sent_to = replace(Request.Form("Registration_sent_to"),"'","''")
			order_notification_subject = replace(Request.Form("order_notification_subject"),"'","''")
			order_notification_body = replace(Request.Form("order_notification_body"),"'","''")
			order_notification_sent_to = replace(Request.Form("order_notification_sent_to"),"'","''")
			Shipping_email_subject = replace(Request.Form("Shipping_email_subject"),"'","''")
			Shipping_email_body = replace(Request.Form("Shipping_email_body"),"'","''")
			Newsletter_notification_subject = replace(Request.Form("Newsletter_notification_subject"),"'","''")
               Newsletter_notification = replace(Request.Form("Newsletter_notification"),"'","''")

			Gift_email_Subject = replace(Request.Form("Gift_email_Subject"),"'","''")
			Gift_buyer_email_body = replace(Request.Form("Gift_buyer_email_body"),"'","''")

			pin_email_Subject = replace(Request.Form("pin_email_Subject"),"'","''")
			pin_buyer_email_body = replace(Request.Form("pin_buyer_email_body"),"'","''")

			sql_update = "UPDATE Store_emails SET Newsletter_notification_subject='"&Newsletter_notification_subject&"', Newsletter_notification='"&Newsletter_notification&"',pin_buyer_email_body='" &pin_buyer_email_body& "',pin_email_Subject='"&pin_email_Subject&"', Gift_buyer_email_body='"&Gift_buyer_email_body&"', Gift_email_Subject='"&Gift_email_Subject&"', Registration_enable = "&Registration_enable&", Registration_user_psw = "&Registration_user_psw&", Registration_subject = '"&Registration_subject&"', Registration_body = '"&Registration_body&"', Registration_sent_to = '"&Registration_sent_to&"', order_notification_to_admin_enable = "&order_notification_to_admin_enable&", order_notification_enable = "&order_notification_enable&", order_notification_subject = '"&order_notification_subject&"', order_notification_body = '"&order_notification_body&"', order_notification_invoice = "&order_notification_invoice&", order_notification_sent_to = '"&order_notification_sent_to&"', Shipping_email_subject = '"&Shipping_email_subject&"', Shipping_email_body = '"&Shipping_email_body&"' WHERE Store_id="&Store_id
			
			conn_store.Execute sql_update
            sql_update = "UPDATE Store_settings SET Html_Notifications="&Html_Notifications&" WHERE Store_id="&Store_id
            conn_store.Execute sql_update
            

		'CAME FROM ADD STORE LOGINS WINDOW
		case "Security_Add"
			Access_Priviledge = Request.Form("Access_Priviledge")
				Store_User_id = checkStringForQ(Request.Form("Store_User_id"))
				Store_Password = checkStringForQ(Request.Form("Store_Password"))
				Store_First_Name = checkStringForQ(Request.Form("Store_First_Name"))
				Store_Last_Name = checkStringForQ(Request.Form("Store_Last_Name"))
				Store_Email = checkStringForQ(Request.Form("Store_Email"))
	
				'CHECH IF USER ID ALREADY EXISTS
				sql_check = "select Login_Id from store_logins WITH (NOLOCK) where Store_User_ID = '"&Store_User_id&"' and store_id = "&store_id
				rs_Store.open sql_check,conn_store,1,1 
				if rs_Store.bof = false then
					rs_Store.close
					response.redirect "admin_error.asp?Message_id=8"
				end if
				rs_Store.close
				if Request.Form("Store_Password") <> Request.Form("Store_Password_Confirm") then
					response.redirect "admin_error.asp?Message_id=10"
				end if
				sql_insert = "Insert into store_logins (Store_id,Store_User_id,Store_Password,Store_First_Name,Store_Last_Name,Store_Email,Login_Privilege,Access_Priviledge) values ("&Store_id&",'"&Store_User_id&"','"&Store_Password&"','"&Store_First_Name&"','"&Store_Last_Name&"','"&Store_Email&"',2,'"&Access_Priviledge&"')"
				conn_store.Execute sql_insert

		'CAME FROM EDIT STORE LOGINS WINDOW
		case "Security_Edit"
			Login_Id = Request.Form("Login_Id")
			Access_Priviledge = Request.Form("Access_Priviledge")
			if isNumeric(Login_Id) then
				Store_User_id = checkStringForQ(Request.Form("Store_User_id"))
				Store_Password = checkStringForQ(Request.Form("Store_Password"))
				Store_First_Name = checkStringForQ(Request.Form("Store_First_Name"))
				Store_Last_Name = checkStringForQ(Request.Form("Store_Last_Name"))
				Store_Email = checkStringForQ(Request.Form("Store_Email"))

				'CHECK IF USER ID ALREADY EXISTS
				sql_check = "select Login_Id from store_logins WITH (NOLOCK) where Store_User_ID = '"&Store_User_id&"' and login_id <> "&Login_Id&" and store_id = "&store_id
				rs_Store.open sql_check,conn_store,1,1 
				if rs_Store.bof = false then
					rs_Store.close
					response.redirect "admin_error.asp?Message_id=8"
				end if
				rs_Store.close
				if Request.Form("Store_Password") <> Request.Form("Store_Password_Confirm") then
					response.redirect "admin_error.asp?Message_id=10"
				end if
				sql_update = "update store_logins set Store_User_id='"&Store_User_id&"', Store_Password = '"&Store_Password&"', Store_First_Name = '"&Store_First_Name&"', Store_Last_Name = '"&Store_Last_Name&"', Store_Email = '"&Store_Email&"', Access_Priviledge = '"&Access_Priviledge&"' where Store_id = "&Store_id&" and login_id = "&Login_ID
				conn_store.Execute sql_update
			else
				Response.Redirect "error.asp?message_id=1"
			end if

		'CAME FROM STORE MANAGEMENT (START VALUES FOR OID/IID/CID/TID) WINDOW
		case "Store_Manag_Starts"
			  StartOID = clng(request.form("StartOID"))
			StartCID = clng(request.form("StartCID"))
			if isNumeric(StartOID) and isNumeric(CID) then
			sql_update = "update store_settings set StartOID="&StartOID&", StartCID="&StartCID&" where store_id="&store_id
			conn_store.execute sql_update
		else
				 Response.Redirect "admin_error.asp?message_id=1"
		end if
		case "Shipping_Labels"
			Page_Width = request.form("Page_Width")
			Page_Height = request.form("Page_Height")
			Address_Width = request.form("Address_Width")
			Address_Height = request.form("Address_Height")
			Spacing_Width = request.form("Spacing_Width")
			Spacing_Height = request.form("Spacing_Height")
			Top_Margin = request.form("Top_Margin")
			Bottom_Margin = request.form("Bottom_Margin")
			Left_Margin = request.form("Left_Margin")
			Right_Margin = request.form("Right_Margin")
			Total_Rows = request.form("Total_Rows")
			Total_Cols = request.form("Total_Cols")
			Image_Name = replace(request.form("Image_Name"),"'","''")
			Image_Pos = request.form("Image_Pos")
			if isNumeric(Page_Width) and isNumeric(Page_Height) and isNumeric(Address_Width) and isNumeric(Address_Height) and isNumeric(Spacing_Width) and isNumeric(Spacing_Height) and isNumeric(Top_Margin) and isNumeric(Bottom_Margin) and isNumeric(Left_Margin) and isNumeric(Right_Margin) and isNumeric(Total_Rows) and isNumeric(Total_Cols) then
				sql_update = "update Store_Shipping_Labels set Image_Name='"&Image_Name&"', Image_Pos="&Image_Pos&", Page_Width="&Page_Width&", Page_Height="&Page_Height&", Address_Width="&Address_Width&", Address_Height="&Address_Height&", Spacing_Width="&Spacing_Width&", Spacing_Height="&Spacing_Height&", Top_Margin="&Top_Margin&", Bottom_Margin="&Bottom_Margin&", Left_Margin="&Left_Margin&", Right_Margin="&Right_Margin&", Total_Rows="&Total_Rows&", Total_Cols="&Total_Cols&" where store_id="&store_id
				conn_store.execute sql_update
			 else
				Response.Redirect "admin_error.asp?message_id=1"
			 end if

		case "Froogle"
			  Froogle_Filename = checkStringForQ(request.form("Froogle_Filename"))
				Froogle_Username = checkStringForQ(request.form("Froogle_Username"))
				Froogle_Password = checkStringForQ(request.form("Froogle_Password"))
				Froogle_Server = checkStringForQ(request.form("Froogle_Server"))
				Froogle_Enable = checkStringForQ(request.form("Froogle_Enable"))
				Force_Resend = checkStringForQ(request.form("Force_Resend"))
				Large_Upload = checkStringForQ(request.form("Large_Upload"))
		
				if Froogle_Enable <> "" then
					Froogle_Enable = 1
				else
					Froogle_Enable = 0
				end if
				if Force_Resend <> "" then
					Force_Resend = 1
				else
					Force_Resend = 0
				end if
				sql_update = "update store_External set Force_Resend="&Force_Resend&",Large_Upload="&Large_Upload&",Froggle_Filename='"&Froogle_Filename&"', Froggle_Username='"&Froogle_Username&"', Froggle_Password='"&Froogle_Password&"', Froggle_Server='"&Froogle_Server&"', Froggle_Enable="&Froogle_Enable&" where store_id="&store_id
				conn_store.execute sql_update

		case "homepage" 
			sHomepage = nullifyQ(Request.Form("sHomepage"))
			sql_update_instance = "exec wsp_page_content_top "&Store_id&",5,'"&sHomepage&"';"
			conn_store.Execute sql_update_instance
			Store_Homepage = lcase(request.form("Store_Homepage"))
			if Store_Homepage<>"" then
               if Store_Homepage="http://"&Store_Domain or Store_Homepage="http://www."&Store_Domain or Store_Homepage="http://"&Store_Domain2 or Store_Homepage="http://www."&Store_Domain2 or Store_Homepage=Site_Name2 or instr(Store_Homepage,"store.asp")>0 then
  			         response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("This page is already your homepage.  Please leave the homepage field blank to use the existing store.asp as the homepage.")
  			      elseif instr(Store_Homepage,"http://")=0 and instr(Store_Homepage,"https://")=0  then
                  response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("The default page url you specified does not appear valid.  The page must begin with http:// or https://")
               end if
         end if

			sql_update = "update store_settings set Store_Homepage='"&Store_Homepage&"' where store_id="&store_id
			conn_store.execute sql_update
		case "returns"
			Return_Policy = nullifyQ(Request.Form("Return_Policy"))
			sql_update_instance = "exec wsp_page_content_top "&Store_id&",14,'"&Return_Policy&"';"
			conn_store.Execute sql_update_instance
			Auto_RMA_Days = Request.Form("Auto_RMA_Days")
			if not isNumeric(Auto_RMA_Days) then
				Auto_RMA_Days = 0
			end if

			Enable_RMA = Request.Form("Enable_RMA")
			if Enable_RMA = "" then
				Enable_RMA = 0
			else
				Enable_RMA = 1
			end if
			sql_update = "update store_settings set enable_rma="&enable_rma&", auto_rma_days="&auto_rma_days&" where store_id="&store_id
			conn_store.execute sql_update

		 case "privacy"
			Privacy_Policy = nullifyQ(Request.Form("Privacy_Policy"))

			sql_update_instance = "exec wsp_page_content_top "&Store_id&",18,'"&Privacy_Policy&"';"
			conn_store.Execute sql_update_instance
         
			
    case "Store_Dept_Layout"
			Department_Layout_Id = Request.Form("Department_Layout_Id")
			if Department_Layout_Id = 6 then
				Department_Layout = replace(Request.Form("Department_Layout"),"'","''")
			else
				sql_select = "select Template from Sys_Inventory_Template where Template_Type=1 and Template_Id="&Department_Layout_Id
				rs_Store.open sql_select,conn_store,1,1
				if not rs_Store.bof and not rs_Store.eof then
					Department_Layout = replace(rs_Store("Template"),"'","''")
				else
					Response.Redirect "admin_error.asp?message_id=1"
				end if
			end if 
			sql_update = "UPDATE Store_Settings SET Department_Layout = '"&Department_Layout&"',Department_Layout_Id = "&Department_Layout_Id&" WHERE Store_id="&Store_id&";"
			conn_store.Execute sql_update
			

		case "Store_Item_Layout"

			Item_S_Layout_Id = Request.Form("Item_S_Layout_Id")
			Item_L_Layout_Id = Request.Form("Item_L_Layout_Id")
			Item_F_Layout_Id = Request.Form("Item_F_Layout_Id")

			if Item_S_Layout_Id = 11 then
				Item_S_Layout = replace(Request.Form("Item_S_Layout"),"'","''")
			else 
				sql_select = "select Template from Sys_Inventory_Template where Template_Type=2 and Template_Id="&Item_S_Layout_Id
				rs_Store.open sql_select,conn_store,1,1
				
				if not rs_Store.bof and not rs_Store.eof then
					
					Item_S_Layout = replace(rs_Store("Template"),"'","''")
				else
					Response.Redirect "admin_error.asp?message_id=1"
				end if
				rs_Store.close
			end if

			if Item_F_Layout_Id = 11 then
				Item_F_Layout = replace(Request.Form("Item_F_Layout"),"'","''")
			else
				sql_select = "select Template from Sys_Inventory_Template where Template_Type=2 and Template_Id="&Item_F_Layout_Id
				rs_Store.open sql_select,conn_store,1,1

				if not rs_Store.bof and not rs_Store.eof then

					Item_F_Layout = replace(rs_Store("Template"),"'","''")
				else
					Response.Redirect "admin_error.asp?message_id=1"
				end if
				rs_Store.close
			end if

			if Item_L_Layout_Id = 11 then
				Item_L_Layout = replace(Request.Form("Item_L_Layout"),"'","''")
			else 
				sql_select = "select Template from Sys_Inventory_Template where Template_Type=2 and Template_Id="&Item_L_Layout_Id
				rs_Store.open sql_select,conn_store,1,1
				if not rs_Store.bof and not rs_Store.eof then
					
					Item_L_Layout = replace(rs_Store("Template"),"'","''")
				else
					Response.Redirect "admin_error.asp?message_id=1"
				end if
				rs_Store.close
			
			end if
			if Item_F_Layout_Id="" then
				Item_F_Layout_Id=6
			end if
			if Item_S_Layout_Id="" then
				Item_S_Layout_Id=6
			end if
			if Item_L_Layout_Id="" then
				Item_L_Layout_Id=6
			end if
			on error goto 0
			sql_update = "UPDATE Store_Settings SET Item_F_Layout = '"&Item_F_Layout&"',Item_S_Layout = '"&Item_S_Layout&"',Item_L_Layout = '"&Item_L_Layout&"',Item_F_Layout_Id = "&Item_F_Layout_Id&",Item_S_Layout_Id = "&Item_S_Layout_Id&",Item_L_Layout_Id = "&Item_L_Layout_Id&" WHERE Store_id="&Store_id&";"
			conn_store.Execute sql_update

  Case "statistics"
			GMT_Offset = Request.Form("GMT_Offset")
			Skip_Hosts = Request.Form("Skip_Hosts")
			DNS_Lookup = Request.Form("DNS_Lookup")
			if DNS_Lookup = "" then
			   DNS_Lookup = 0
			end if
      sql_Update = "Update Store_Statistics set DNS_Lookup="&DNS_Lookup&", GMT_Offset = "&GMT_Offset&", Skip_Hosts = '"&Skip_Hosts&"' where Store_id = "&Store_id
			conn_store.Execute sql_Update

		case "preCharge"
				preCharge_MerchantID = checkStringForQ(request.form("preCharge_MerchantID"))
				preCharge_Security1 = checkStringForQ(request.form("preCharge_Security1"))
				preCharge_Security2 = checkStringForQ(request.form("preCharge_Security2"))
				preCharge_Enable = checkStringForQ(request.form("preCharge_Enable"))
				if preCharge_Enable <> "" then
					preCharge_Enable = 1
				else
					preCharge_Enable = 0
				end if
				sql_update = "update store_external set  Precharge_MerchantID='"&preCharge_MerchantID&"', Precharge_Security1='"&preCharge_Security1&"', Precharge_Security2='"&preCharge_Security2&"',Precharge_Enable="&preCharge_Enable&" where store_id="&store_id

  			conn_store.execute sql_update
		case "maxmind"
				maxmind_license = checkStringForQ(request.form("maxmind_license"))
				maxmind_reject = checkStringForQ(request.form("maxmind_reject"))
				maxmind_Enable = checkStringForQ(request.form("maxmind_Enable"))
				if maxmind_Enable <> "" then
					maxmind_Enable = 1
				else
					maxmind_Enable = 0
				end if
				sql_update = "update store_external set  maxmind_license='"&maxmind_license&"', maxmind_reject="&maxmind_reject&", Maxmind_Enable="&Maxmind_Enable&" where store_id="&store_id

  			conn_store.execute sql_update

	end select

	'CAME FROM DELETE STORE LOGINS WINDOW
	if Request.QueryString("Delete_Login_Id") <> "" then
		Login_Id = Request.QueryString("Delete_Login_Id")
		sql_delete = "Delete from store_logins where Login_Id = "&Login_Id
		conn_store.Execute sql_delete

	end if

	if err.number <> 0 then
		if instr(err.description,"String or binary data would be truncated") then
			sError = "You have entered to much data please shorten 1 or more fields and try again.  The limit for a single record is 8060 bytes which you have exceeded."
		else
			sError = err.description
		end if

		fn_error sError
  end if

	response.redirect returnTo

end if

response.write returnTo

%>

