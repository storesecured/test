<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->
<!--#include file="help/message_customers.asp"-->

<%

Server.ScriptTimeout =500

if inStr(Store_Domain,"www.") > 0 then
	  iLength = len(Store_Domain)
	  Store_Domain = Right(Store_Domain,iLength-4)
end if

if request.form("Message_Customers") <> "" then
   set myEmails = server.createobject("scripting.dictionary")

if request.form("From")<>"" then
From = request.form("From")
else
From = Session("Store_Email")
end if

	
	StartBody = request.form("Body")
	StartBody = replace(StartBody,"%OBJ_STORE_IMAGES_OBJ%",Site_Name&"/images")
	StartSubject = request.form("Subject")

	Customer_Group = request.form("Customer_Group")
        if Customer_Group = "ToMerchant" then
        	
                ToEmail = lcase(request.form("ToMerchant"))
                Body = StartBody
		Subject = StartSubject
                Send_Mail_Html From,ToEmail,Subject,Body
        else
	if Customer_Group = "Newsletter" then
		sql_customers = "select * from Store_Newsletter WITH (NOLOCK) where Store_Id="&Store_Id

		set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_customers,mydata,myfields,noRecords)
		if noRecords = 0 then
		 FOR rowcounter= 0 TO myfields("rowcount")
			ToEmail = lcase(mydata(myfields("email_address"),rowcounter))
			if not(myEmails.Exists(ToEmail)) then
			   myEmails.Add ToEmail, "1"
                           if InStr(StartBody,"%") > -1 then
				Body = Replace(StartBody,"%LASTNAME%",mydata(myfields("last_name"),rowcounter))
				Body = Replace(Body,"%FIRSTNAME%",mydata(myfields("first_name"),rowcounter))
			else
				Body = StartBody
			end if
			if InStr(StartSubject,"%") > -1 then
				Subject = Replace(StartSubject,"%LASTNAME%",mydata(myfields("last_name"),rowcounter))
				Subject = Replace(Subject,"%FIRSTNAME%",mydata(myfields("first_name"),rowcounter))
			else
				Subject = StartSubject
			end if
			sURL = Site_Name&"cancel_signup.asp?page_id=40&Store_id="&Store_Id&"&Email_Address="&Server.urlEncode(ToEmail)
			Body = Body & "<BR><BR><font size=1>" & "If you no longer wish to receive our newsletter <a href='"&sURL&"'>click here</a> to be removed.</font>"
                        
                           Send_Mail_Html From,ToEmail,Subject,Body
                        end if
		Next
		end if


	else


		if isNumeric(Customer_Group) then
			sql_groups = "select * from Store_Customers_Groups WITH (NOLOCK) where Store_id = "&Store_id&" and Group_id="&Customer_Group
			set myfields=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_groups,mydata,myfields,noRecords)

			group_country = mydata(myfields("group_country"),rowcounter)
			if UCase(group_country) = "ALL" then
				Country_String = ""
			else
				group_country_C = Replace(group_country,",","','")
				Country_String = "country in ('"&group_country_C&"') and "
			end if

			  group_dept = mydata(myfields("group_dept"),rowcounter)
			  if UCase(group_dept) = "ALL" then
				Dept_String = ""
			elseif group_dept <> "" then
				group_dept_C = Replace(group_dept,",","','")
				Dept_String = "Orders_Dept in ('"&group_dept_C&"') and "
			else
				  Dept_String = ""
			end if

			group_Company = mydata(myfields("group_company"),rowcounter)
			if UCase(group_Company) = "ALL" then
				Company_String = ""
			else
				group_Company_C = Replace(group_Company,",","','")
				Company_String = "Company in ('"&group_Company_C&"') and "
			end if
		
			group_Cid = mydata(myfields("group_cid"),rowcounter)
			if UCase(group_Cid) = "" then
				Cid_String = "1=1) and"
			else
				group_Cid_C = group_Cid
				Cid_String = "1=1) OR (CCid in ("&group_Cid_C&"))) and ("
			end if
		


			sql_customers = "select * FROM Store_Customers WITH (NOLOCK) Where ((Orders_Total >= "&mydata(myfields("group_purchase_history"),rowcounter)&" and Budget_left >= "&mydata(myfields("group_budget_min"),rowcounter)&" and  "&Country_String&" "&Company_String&" "&Cid_String&" Record_type=0 and Store_id = "&Store_id&" and spam <>0)"
		else
			sql_customers = "select * from Store_Customers WITH (NOLOCK) where Store_Id="&Store_Id&" and record_type=0 and spam<>0"
		end if
		
      set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_customers,mydata,myfields,noRecords)

		if noRecords = 0 then

		FOR rowcounter= 0 TO myfields("rowcount")
			Orders_Dept = mydata(myfields("orders_dept"),rowcounter)
         if (Orders_Dept = "" or isnull(Orders_Dept)) and group_dept <> "" then
				sIncluded = 0
			elseif group_dept = "" or isnull(group_dept) then
				sIncluded = 1
			else
				sIncluded=0

				sArray = split(Orders_Dept, ",")
				for each one in sArray
					 sArrayCurrent = split(group_dept,",")
					 if sIncluded = 0 then
						 for each Dept in sArrayCurrent
							  if Dept=one and sIncluded=0 then
								  sIncluded = 1
								end if
						 next
					  end if
				next
			end if

			if sIncluded=1 then

				ToEmail = lcase(mydata(myfields("email"),rowcounter))
				if not(myEmails.Exists(ToEmail)) then
                                   myEmails.Add ToEmail, "1"
				if InStr(StartBody,"%") > -1 then
					Body = Replace(StartBody,"%LASTNAME%",mydata(myfields("last_name"),rowcounter))
					Body = Replace(Body,"%FIRSTNAME%",mydata(myfields("first_name"),rowcounter))
					Body = Replace(Body,"%CITY%",mydata(myfields("city"),rowcounter))
					Body = Replace(Body,"%LOGIN%",mydata(myfields("user_id"),rowcounter))
					Body = Replace(Body,"%PASSWORD%",mydata(myfields("password"),rowcounter))
					Body = Replace(Body,"%LASTVISIT%",mydata(myfields("lastaccess"),rowcounter))
					Body = Replace(Body,"%BUDGETLEFT%",mydata(myfields("budget_left"),rowcounter))
					Body = Replace(Body,"%REWARDLEFT%",mydata(myfields("reward_left"),rowcounter))
					Body = Replace(Body,"%ORDERSTOTAL%",mydata(myfields("orders_total"),rowcounter))
					Body = Replace(Body,"%FIRSTVISIT%",mydata(myfields("registration_date"),rowcounter))
				else
					Body = StartBody
				end if

				if InStr(StartSubject,"%") > -1 then
					Subject = Replace(StartSubject,"%LASTNAME%",mydata(myfields("last_name"),rowcounter))
					Subject = Replace(Subject,"%FIRSTNAME%",mydata(myfields("first_name"),rowcounter))
				else
					Subject = StartSubject
				end if
				
				sURL = Site_Name&"cancel_signup.asp?page_id=40&Store_id="&Store_Id&"&Email_Address="&Server.urlEncode(ToEmail)
			        Body = Body & "<BR><BR><font size=1>" & "If you no longer wish to receive our promotional emails <a href='"&sURL&"'>click here</a> to be removed.</font>"

				Send_Mail_Html From,ToEmail,Subject,Body
				end if

		  end if
		Next
		end if
	end if
	end if
set myEmails=Nothing
end if
on error resume next
sTextHelp="newsletter/newsletter_send.doc"

sFormAction = "message_customers.asp"
sName = "Message_Customers"
sFormName = "customers"
sSubmitName = "Message_Customers"
thisRedirect = "message_customers.asp"
sTopic="Message_Customers"
sMenu="marketing"
addPicker=1
sTitle = "Send Newsletter"
sFullTitle = "Marketing > <a href=newsletter_manager.asp class=white>Newsletter</a> > Send"
sQuestion_Path = "marketing/mail_merge.htm"
createHead thisRedirect
if Service_Type < 1	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>

<% else %>
    <tr bgcolor='#FFFFFF'>
			<td width="100%" colspan="3" height="15">&nbsp;
			<input type="button" OnClick=JavaScript:self.location="customers_groups.asp" class="Buttons" value="Customer Group List" name="Create_new_Page_link">
			<input type="button" OnClick=JavaScript:self.location="newsletter_manager.asp" class="Buttons" value="Newsletter Subscriber List" name="Create_new_Page_link">
			</td>
		</tr>
		<tr bgcolor='#FFFFFF'><td class="inputname"><b>From</b></td><td class="inputvalue">
     <input type=text name=From value='<%=Store_Email%>' size=38>

		<% small_help "From" %></td></tr>
		<tr bgcolor='#FFFFFF'><td class="inputname"><b>To</b></td>
		<td class="inputvalue"><select name="Customer_Group">
		<option value="All" onClick="this.form.ToMerchant.disabled = true">All Customers who want promo emails</option>
		<option value="Newsletter" onClick="this.form.ToMerchant.disabled = true">Newsletter Subscribers</option>
		<option value="ToMerchant" onClick="this.form.ToMerchant.disabled = false">Test Email Address Only</option>
		<% sql_groups = "select Group_Id,Group_Name from Store_Customers_Groups where Store_id = "&Store_id&""
		rs_Store.open sql_groups,conn_store,1,1

		Do While Not rs_Store.EOF %>
			<option value="<%= rs_Store("Group_Id") %>" onClick="this.form.ToMerchant.disabled = true"><%= rs_Store("Group_Name") %> who want promo emails</option>
		<% rs_store.movenext
		Loop
		rs_Store.Close %>
		</select> <br>
		 <input type=text name=ToMerchant value='' size=38 disabled=true>
		<% small_help "To" %></td></tr>
		<tr bgcolor='#FFFFFF'><td class="inputname"><b>Subject</b></td><td class="inputvalue"><input type=text name=Subject value='' size=38>
		<% small_help "Subject" %></td></tr>
		<tr bgcolor='#FFFFFF'><td colspan=3 class=instructions>To personalize emails use the following special identifiers in email body (note if
		selecting newsletter subscribers only firstname and lastname identifiers will work):<br>
		%FIRSTNAME% %LASTNAME% %CITY% %LOGIN% %PASSWORD%
		%LASTVISIT% %BUDGETLEFT% %REWARDLEFT% %ORDERSTOTAL% %FIRSTVISIT%<br><br>
		%FIRSTNAME% %LASTNAME% may also be used in the subject line</td></tr>
		<tr bgcolor='#FFFFFF'><td class="inputname" colspan=2><B>Body</b><BR>
		<% on error goto 0 %>
		<%= create_editor ("Body","","[""First Name"",""%FIRSTNAME%""],[""Last Name"",""%LASTNAME%""],[""Login"",""%LOGIN%""],[""Password"",""%PASSWORD%""],[""City"",""%CITY%""],[""Last Visit"",""%LASTVISIT%""],[""Budget Left"",""%BUDGETLEFT%""],[""Reward Left"",""%REWARDLEFT%""],[""Order Total"",""%ORDERSTOTAL%""],[""First Visit"",""%FIRSTVISIT%""]") %>
		<% small_help "Body" %></td></tr>
		<tr bgcolor='#FFFFFF'><td colspan=3 align=center><input class=buttons type=submit name=Message_Customers value="Send Newsletter">
                <input type="button" OnClick=JavaScript:self.location="newsletter_manager.asp" class="Buttons" value="Cancel" name="Create_new_Page_link">
			</td></tr>
		<tr bgcolor='#FFFFFF'><td colspan=3 class=instructions>Please note that depending on the size of your customer list this may take a while, please only hit send once.</td></tr>
<% end if %>
<% createFoot thisRedirect, 0

%>


