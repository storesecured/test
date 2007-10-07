<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
'USER CLICKED ON "QUICK EDIT"
if request.form("Form_Name")= "Quick_Edit_Update" then
	
	If Form_Error_Handler(Request.Form) <> "" then
	   Error_Log = Form_Error_Handler(Request.Form)
	   if Error_Log <> "" then
		%> <!--#include file="Include/Error_Template.asp"--><%
		response.end
		end if
	end if


	ids = request.form("delete_ids")
	id_array =  split(ids,",")
	for i = 0 to ubound(id_array)
		varIndex = trim(cstr( id_array(i)))

		'GET THE VARIABLES PASSED
		'====================
		 ccId = cInt(varIndex)

		 user_Id= request.form("user_id_" & varIndex)
		 password=request.form("password_" & varIndex)
		 budget_orig=request.form("budget_orig_" & varIndex)
		 budget_left=request.form("budget_left_" & varIndex)

		 spam=request.form("spam_" & varIndex)
		 if spam <> "" then
			spam =1
		 else
			spam =0
		 end if
		 
		 
		 tax_exempt=request.form("Tax_Exempt_" & varIndex)
		 if tax_exempt <> "" then
			tax_exempt =1
		 else
			tax_exempt = 0
		 end if
		 
		  protected_page_access=request.form("protected_page_access_" & varIndex)
		  if protected_page_access <> "" then
			protected_page_access =1
		  else
			protected_page_access = 0
		  end if

		 first_name = request.form("first_name_" & varIndex)
		 last_name = request.form("last_name_" & varIndex)
		 address1=request.form("address1_" & varIndex)
		 address2=request.form("address2_" & varIndex)
		 city=request.form("city_" & varIndex)
		 state=request.form("state_" & varIndex)
		 country=request.form("country_" & varIndex)
		 zip=request.form("zip_" & varIndex)
		 phone=request.form("phone_" & varIndex)
		 email=request.form("email_" & varIndex)
		 fax=request.form("fax_" & varIndex)

		'PREPARE THE QUERY
		qry = " update  store_customers set user_Id = '" & user_id & "' ,   password ='" &  password & "' , budget_orig =  " &  budget_orig & ",  budget_left=" & budget_left & ", spam= " &  spam & ", tax_exempt= "  & tax_exempt & " ,  protected_page_access = " & protected_page_access & " ,   first_name = '" & first_name  & "' , last_name = '" &  last_name & "' , address1 = '" & address1 & "' ,  address2 = '" & address2 & "' ,  city = '" & city & "' ,  State = '" & State & "',  country = '" & country & "' ,  zip = '" & zip  & "' ,  email = '" &  email & "', phone = '" & phone &"' ,  fax = '" & fax  & "' where ccId = " & varIndex & " and store_id = " & store_id
		

			conn_store.Execute qry
	next
	
end if
response.redirect "my_customer_base.asp"
%>
