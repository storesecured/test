<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
server.scripttimeout = 2000

End_Date = FormatDateTime(now(),2)
sQuestion_Path = "reports/customers.htm"

if request.form="" and request.querystring("Form") <> "" then
	Form_Array = split(request.querystring("Form"),"^")
	col1 = Form_Array(0)
	col1_value = Form_Array(1)
	col1_oper = Form_Array(2)
	col2 = Form_Array(3)
	col2_value = Form_Array(4)
	cstatus = Form_Array(5)
end if

if request.form<>"" then
	if request.form("col1")<>"" then
		col1=request.form("col1")
		col1_value=checkstringforQ(request.form("col1_value"))
		if request.form("col2")<>"" then
			col1_oper=request.form("col1_oper")
			col2=request.form("col2")
			col2_value=checkstringforQ(request.form("col2_value"))
		end if
	end if
end if

if col1<>"" then
	if col1<>"" then
		col1_new=replace(col1,"**",col1_value)
		if instr(col1_new,"<>0")>0 then
			'bit field
			col1_value=lcase(col1_value)
			if col1_value="y" or col1_value="yes" or col1_value="1" or col1_value="t" or col1_value="true" then
			else
				col1_new=replace(col1_new,"<>0","=0")
			end if
		elseif instr(col1_new,"'")=0 then
			'no quotes means it should be a number, check
			if not isNumeric(col1_value) then
				fn_error "The search has failed.  You have entered a non-numeric value ("&col1_value&") into a field that requires a number."
			end if
		elseif instr(col1_new,"created")>0 or instr(col1_new,"date")>0 then
			'keywords means its a date
			if not isDate(col1_value) then
				fn_error "The search has failed.  You have entered a non-date value ("&col1_value&") into a field that requires a date."
			else
             col1_value=formatdatetime(col1_value)
             col1_new=replace(col1,"**",col1_value)
         end if
		end if
		sql_where=sql_where & " ("&col1_new&") "
		if col2<>"" then
			col2_new=replace(col2,"**",col2_value)
			if instr(col2_new,"<>0")>0 then
				'bit field
				col2_value=lcase(col2_value)
				if col2_value="y" or col2_value="yes" or col2_value="1" or col2_value="t" or col2_value="true" then
				else
					col2_new=replace(col2_new,"<>0","=0")
				end if
			elseif instr(col2_new,"'")=0 then
				'no quotes means it should be a number, check
				if not isNumeric(col2_value) then
					fn_error "The search has failed.  You have entered a non-numeric value ("&col2_value&") into a field that requires a number."
				end if
			elseif instr(col2_new,"created")>0 or instr(col2_new,"date")>0 then
				'keywords means its a date
				if not isDate(col2_value) then
					fn_error "The search has failed.  You have entered a non-date value ("&col2_value&") into a field that requires a date."
				else
                 col2_value=formatdatetime(col2_value)
                 col2_new=replace(col2,"**",col2_value)
    			end if
    		end if
			sql_where=sql_where & col1_oper & " ("&col2_new&") "
		end if
		sql_where = "("&sql_where&") AND "
	end if
end if


if Request.form("cmbStatus") = "1" then
	cstatus = "1"
	sql_where_end = " and ccid in (select ccid from store_purchases WITH (NOLOCK) where Store_id = "&store_id&" and purchase_completed<>0)"
elseif Request.form("cmbStatus") = "2" then
	cstatus = "2"
	sql_where_end = " and ccid not in (select ccid from store_purchases WITH (NOLOCK) where Store_id = "&store_id&" and purchase_completed<>0)"
else
	cstatus = "0"
	sql_where_end = ""
end if


sSubmitName="SearchCustomers"
End_Date1 = DateAdd("d",1,End_Date)
sNeedSubmit=1
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_customers"
myStructure("TableWhere") = sql_where &" Record_type=0 "& sql_where_end
myStructure("ColumnList") = "ccid,sys_created,last_name,first_name,email"
myStructure("HeadingList") = "ccid,sys_created,last_name,first_name,email"
myStructure("DefaultSort") = "last_name"
myStructure("PrimaryKey") = "ccid"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Footer") = "<Input class=buttons type='Button' value='Mass Edit Selected' onclick=""JavaScript:goMassEdit();"">&nbsp;<Input class=buttons type='Button' value='Quick Edit Selected' onclick=""JavaScript:goQuickEdit();"">"
myStructure("ViewBudget") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "customers"
myStructure("FileName") = "my_customer_base.asp"
myStructure("FormAction") = "my_customer_base.asp"
myStructure("Title") = "Customer Search"
myStructure("FullTitle") = "Customers > Search"
myStructure("CommonName") = "Customer"
myStructure("NewRecord") = "Add_New_Customer.asp"
myStructure("EditRecord") = "modify_customer.asp"
myStructure("viewBudgetRecord") = "budget_view.asp"
myStructure("Heading:ccid") = "PK"
myStructure("Heading:ccid") = "ID"
myStructure("Heading:sys_created") = "Date"
myStructure("Heading:last_name") = "Last Name"
myStructure("Heading:first_name") = "First Name"
myStructure("Heading:email") = "Email"
myStructure("Format:ccid") = "STRING"
myStructure("Format:sys_created") = "DATE"
myStructure("Format:last_name") = "STRING"
myStructure("Format:first_name") = "STRING"
myStructure("Format:email") = "STRING"
myStructure("Length:email") = "25"

myStructure("Link:email") = "mailto:THISFIELD"
myStructure("Form") = col1&"^"&col1_value&"^"&col1_oper&"^"&col2&"^"&col2_value&"^"&cStatus

File_Name = "customers_on_"&End_Date&".txt"
File_Name = Replace(File_Name,"/","-")

if Request.Form("Export_To_Text_File") <> "" then
	sql_select_customers = "select * from "&myStructure("TableName")&" WITH (NOLOCK) where "&sql_where &" store_id="&Store_id&" and (record_type=0 or record_type=1)" & sql_where_end & " order by cid,record_type"
    set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select_customers,mydata,myfields,noRecords)
	TxtExport="User ID" & chr(9) & "Password" & chr(9) & "Billing Last Name" & chr(9) &_
		"Billing First Name" & chr(9) & "Billing Company" & chr(9) & "Billing Address1" & chr(9) &_
		"Billing Address2" & chr(9) & "Billing City" & chr(9) & "Billing State" & chr(9) &_
		"Billing Zip" & chr(9) & "Billing Country" & chr(9) & "Billing Phone" & chr(9) &_
		"Billing Fax" & chr(9) & "Billing Email" & chr(9) & "Shipping Last Name" & chr(9) &_
		"Shipping First Name" & chr(9) & "Shipping Company" & chr(9) & "Shipping Address1" & chr(9) &_
		"Shipping Address2" & chr(9) & "Shipping City" & chr(9) & "Shipping State" & chr(9) &_
		"Shipping Zip" & chr(9) & "Shipping Country" & chr(9) & "Shipping Phone" & chr(9) &_
		"Shipping Fax" & chr(9) & "Shipping Email" & chr(9) & "Promotional Emails" & chr(9) &_
		"Tax Exempt" & chr(9) & "Shipping Same" & chr(9) & "Protected Page Access" & chr(9) &_
        "Budget Original" & chr(9) & "Budget Left" & chr(9) & "Reward Total" & chr(9) &_
		"Reward Left" & chr(9) & chr(13) & chr(10)
	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			If Response.IsClientConnected=false Then
				exit for
			end if
			if rowcounter mod 500 = 0 then
			        'every 500 lines pause for 5 seconds, to prevent from using to much cpu     
                                Set WaitObj = Server.CreateObject ("WaitFor.Comp")
                                WaitObj.WaitForSeconds 5
                                set WaitObj=Nothing
                
                        end if
			if mydata(myfields("record_type"),rowcounter)=0 then
				TxtExportCust=mydata(myfields("user_id"),rowcounter) & chr(9) &_
				mydata(myfields("password"),rowcounter) & chr(9)
			else
			   TxtExportCust=""
			end if
			TxtExportCust=TxtExportCust & (mydata(myfields("last_name"),rowcounter) & chr(9) &_
				mydata(myfields("first_name"),rowcounter) & chr(9) &_
				mydata(myfields("company"),rowcounter) & chr(9) &_
				mydata(myfields("address1"),rowcounter) & chr(9) &_
				mydata(myfields("address2"),rowcounter) & chr(9) &_
				mydata(myfields("city"),rowcounter) & chr(9) &_
				mydata(myfields("state"),rowcounter) & chr(9) &_
				mydata(myfields("zip"),rowcounter) & chr(9) &_
				mydata(myfields("country"),rowcounter) & chr(9) &_
				mydata(myfields("phone"),rowcounter) & chr(9) &_
				mydata(myfields("fax"),rowcounter) & chr(9) &_
				mydata(myfields("email"),rowcounter) & chr(9))
			if mydata(myfields("record_type"),rowcounter)=0 then
				if mydata(myfields("spam"),rowcounter) then
					Spam="Y"
				else
					Spam="N"
				end if
				if mydata(myfields("tax_exempt"),rowcounter) then
					Tax_Exempt="Y"
				else
					Tax_Exempt="N"
				end if
				if mydata(myfields("protected_page_access"),rowcounter) then
					Protected_access = "Y"
				else
					Protected_access = "N"
				end if
				Budget_orig = mydata(myfields("budget_orig"),rowcounter)
				Budget_left = mydata(myfields("budget_left"),rowcounter)
				Reward_Total = mydata(myfields("reward_total"),rowcounter)
				Reward_Left = mydata(myfields("reward_left"),rowcounter)
			end if
			if mydata(myfields("record_type"),rowcounter)=1 then
    			TxtExportCust=TxtExportCust & Spam & chr(9) &_
				Tax_Exempt & chr(9) & "N" & chr(9) & Protected_access & chr(9) &_
				Budget_orig & chr(9) & Budget_left & chr(9) & Reward_Total & chr(9) &_
				Reward_left & chr(9) & chr(13) & chr(10)
			end if

			TxtExport=TxtExport&TxtExportCust
		Next
	end if
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	Export_Folder = fn_get_sites_folder(Store_Id,"Export")
	File_Full_Name = Export_Folder&"\"&File_Name
	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
	MyFile.Write TxtExport
	MyFile.Close
	if Request.Form("Export_To_QuickBooks_File") = "" then
		Response.Redirect "Download_Exported_Files.asp"
	end if
end if 

if Request.Form("Export_To_QuickBooks_File") <> "" then

	'added 21 june 2007 for selected customers.
	
	selected_customers=  request.form("DELETE_IDS")
	
	'added completed

	sql_select_customers = "select * from "&myStructure("TableName")&" WITH (NOLOCK) where "&myStructure("TableWhere")&" and store_id="&Store_id& sql_where_end
	'CODE FOR SAVING THE RESULTS IN A QUICKBOOKS FORMAT FILE
	'FOR USING IN QUICKBOOKS SOFTWARE
	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select_customers,mydata,myfields,noRecords)
	TxtExport = TxtExport&("!CUST"& chr(9) &"NAME"& chr(9) &"REFNUM"& chr(9) &_
		"TIMESTAMP"& chr(9) &"BADDR1"& chr(9) &"BADDR2"& chr(9) &"BADDR3"& chr(9) &_
		"BADDR4"& chr(9) &"BADDR5"& chr(9) &"SADDR1"& chr(9) &"SADDR2"& chr(9) &_
		"SADDR3"& chr(9) &"SADDR4"& chr(9) &"SADDR5"& chr(9) &"PHONE1"& chr(9) &_
		"PHONE2"& chr(9) &"FAXNUM"& chr(9) &"EMAIL"& chr(9) &"NOTE"& chr(9) &_
		"CONT1"& chr(9) &"CONT2"& chr(9) &"CTYPE"& chr(9) &"TERMS"& chr(9) &_
		"TAXABLE"& chr(9) &"LIMIT"& chr(9)&"RESALENUM"& chr(9)&"REP"& chr(9) &_
		"TAXITEM"& chr(9)&"NOTEPAD"& chr(9)&"SALUTATION"& chr(9)&"COMPANYNAME"& chr(9) &_
		"FIRSTNAME"& chr(9)&"MIDINIT"& chr(9)&"LASTNAME"& chr(9)&"CUSTFLD1"& chr(9) &_
		"CUSTFLD2"& chr(9)&"CUSTFLD3"& chr(9)&"CUSTFLD4"& chr(9)&"CUSTFLD5"& chr(9) &_
		"CUSTFLD6"& chr(9)&"CUSTFLD7"& chr(9)&"CUSTFLD8"& chr(9)&"CUSTFLD9"& chr(9) &_
		"CUSTFLD10"& chr(9)&"CUSTFLD11"& chr(9)&"CUSTFLD12"& chr(9)&"CUSTFLD13"& chr(9) &_
		"CUSTFLD14"& chr(9)&"CUSTFLD15"& chr(9)&"JOBDESC"& chr(9)&"JOBTYPE"& chr(9) &_
		"JOBSTATUS"& chr(9)&"JOBSTART"& chr(9)&"JOBPROJEND"& chr(9)&"JOBEND"& chr(9) &_
		"HIDDEN"& chr(9)&"DELCOUNT"& chr(13) & chr(10))
	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			If Response.IsClientConnected=false Then
				exit for
			end if
			if rowcounter mod 500 = 0 then
			        'every 500 lines pause for 5 seconds, to prevent from using to much cpu     
                                Set WaitObj = Server.CreateObject ("WaitFor.Comp")
                                WaitObj.WaitForSeconds 5
                                set WaitObj=Nothing
                
                        end if
			TxtExport = TxtExport&("CUST"& chr(9)&mydata(myfields("first_name"),rowcounter)&_
				" "&mydata(myfields("last_name"),rowcounter)& chr(9)&_
				mydata(myfields("ccid"),rowcounter)& chr(9)&_
				mydata(myfields("registration_date"),rowcounter)& chr(9)&_
				mydata(myfields("first_name"),rowcounter)&_
				" "&mydata(myfields("last_name"),rowcounter)&chr(9)&_
				mydata(myfields("address1"),rowcounter)&_
				" "&mydata(myfields("address2"),rowcounter)&chr(9)&_
				mydata(myfields("city"),rowcounter)&_
				" "&mydata(myfields("state"),rowcounter)&_
				" "&mydata(myfields("zip"),rowcounter)&chr(9)&_
				mydata(myfields("country"),rowcounter)& chr(9)&chr(9))

			sql_sel_shipping = "select * from store_customers WITH (NOLOCK) where record_type<>0 and CID = "&mydata(myfields("cid"),rowcounter)
			set myfields1=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_sel_shipping,mydata1,myfields1,noRecords1)
			ship0=""
			ship1=""
			ship2=""
			ship3=""
			ship4=""
			ship5=""
			if noRecords1 = 0 then
				FOR rowcounter1= 0 TO myfields1("rowcount")
					if ship1="" then
						ship0 = mydata1(myfields1("first_name"),rowcounter1)&" "&mydata1(myfields1("last_name"),rowcounter1)
						ship1 = mydata1(myfields1("address1"),rowcounter1)&" "&mydata1(myfields1("address2"),rowcounter1)
						ship2 = mydata1(myfields1("city"),rowcounter1)&" "&mydata1(myfields1("state"),rowcounter1)&" "&mydata1(myfields1("zip"),rowcounter1)
						ship3 = mydata1(myfields1("country"),rowcounter1)
					end if
				Next
			End If
			TxtExportCust = ship0& chr(9)&ship1& chr(9)&ship2& chr(9)&ship3& chr(9)&ship4& chr(9)
			TxtExportCust = TxtExportCust&(mydata(myfields("phone"),rowcounter)& chr(9)& chr(9)&_
				mydata(myfields("fax"),rowcounter)& chr(9)&mydata(myfields("email"),rowcounter)& chr(9)&_
				chr(9)& chr(9)& chr(9)& chr(9)& chr(9)&"Y"& chr(9)& chr(9)& chr(9)& chr(9)& chr(9)&_
				chr(9)& chr(9)&mydata(myfields("company"),rowcounter)& chr(9)&_
				mydata(myfields("first_name"),rowcounter)& chr(9)& chr(9)&_
				mydata(myfields("last_name"),rowcounter)& chr(9)& chr(9)& chr(9)& chr(9)&_
				chr(9)& chr(9)& chr(9)& chr(9)& chr(9)& chr(9)& chr(9)& chr(9)& chr(9)&_
				chr(9)& chr(9)& chr(9)& chr(9)& chr(9)&"0"& chr(9)& chr(9)& chr(9)& chr(9)&_
				"N"& chr(9)& chr(13) & chr(10))
			TxtExport=TxtExport&TxtExportCust
		Next
	End If
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	File_Name = mid(File_Name,1,len(File_Name)-4)
	File_Name = File_Name&".iif"
	Export_Folder = fn_get_sites_folder(Store_Id,"Export")
	File_Full_Name = Export_Folder&"\"&File_Name
	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
	TxtExport=replace(replace(TxtExport,"&#34;",""""),"&#39;","'")
	MyFile.Write TxtExport
	MyFile.Close 
	'Response.Redirect "Download_Exported_Files.asp"
end if

%>
<script langauge="JavaScript">


  	function goMassEdit()
	{
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked)
				tsel = tsel+1;
		if (tsel<=1){
			alert("To edit a single customer please use the edit link.");
			return;}
		document.forms[0].action="customer_mass_edit.asp";
		document.forms[0].submit();
		
	}

	function goQuickEdit()
	{
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked)
				tsel = tsel+1;
		if (tsel<=1){
			alert("To edit a single customer please use the edit link.");
			return;}
		document.forms[0].action="customer_Quick_edit.asp";
		document.forms[0].submit();
		
	}
</script>


<SCRIPT LANGUAGE="VBScript">
 
	Sub display()	
	

		Dim WshShell
		dim FValue
		dim SValue		
		dim temp
		dim StatusVal
		
		
		Set WshShell = CreateObject("wscript.shell") 	


		temp="<%=sql_select_customers%>"

		if temp="" then
			msgbox "Please Click Export To Quickbooks button First",0,"Quickbooks"
			exit sub
		end if
		
		temp=temp & "|" & "TABLE=CUSTOMERS"   & "|" &"<%=selected_customers%>"	
		
		'msgbox temp
			
	
		
	    WshShell.Run "C:\bin\Quick.exe" & " " & temp  	
		
		

	exit sub



		
		fvalue= "<%=col1_value%>"
		svalue="<%=col2_value%>"
		StatusVal="<%=cstatus%>"		
		
		'Both Text boxes are empty then add only value of status combo as arguments
		If fvalue="" and svalue="" then	
			if StatusVal=0 then
				
			else
				Temp="WHERE " & "made_purchase=" & " " & "<%=cstatus%>"		
			end if			
		'First text box not empty then pass it as argument with Where clause
		ElseIf FValue<>"" and SValue="" then
			if StatusVal=0 then
				Temp="WHERE " & "<%=col1_new%>"
			else
				Temp="WHERE " & "<%=col1_new%>" & " " & "AND" & " " & "made_purchase=" & " " & "<%=cstatus%>"				
			end if			
		'Second text box not empty then pass it as argument with Where clause	
		ElseIf Fvalue="" and SValue<>"" then
			if StatusVal=0 then
				Temp="WHERE " & "<%=col2_new%>"
			else
				Temp="WHERE " & "<%=col2_new%>" & " " & "AND" & " " & "made_purchase=" & " " & "<%=cstatus%>"
			end if			
		Else
		'both text box not empty then pass them as argument with Where clause	
			if StatusVal=0 then
				temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)"
			else
				temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND" & " " & "made_purchase=" & " " & "<%=cstatus%>"		
			end if
			
		end if
		
			temp=temp & "|" & "TABLE=CUSTOMERS"	

		'Run Quick Book Exe	
		WshShell.Run "C:\bin\Quick.exe" & " " & temp
		
	end sub
				
</script>




<!--#include file="head_view.asp"-->

<input type=hidden name=records value=<%= RowsPerPage %> ID="Hidden1">

<TR bgcolor='#FFFFFF'><td colspan=10 width='100%'>

	<table width='100%' border='0' cellspacing='1' cellpadding=2 ID="Table1">
	<tr bgcolor='#FFFFFF'>
	<td class="inputname" width=40%><b>Status</b></td>
	<TD class="inputvalue">
		<SELECT name=cmbStatus ID="Select1">
			<option value = '0' <%if cstatus = 0 then %> selected <% end if%>>All customers</option>
			<option value = '1' <%if cstatus = 1 then %> selected <% end if%>>With purchases</option>
			<option value = '2' <%if cstatus = 2 then %> selected <% end if%>>No purchases</option>
		</SELECT>
	</TD>
	</tr>
	<tr bgcolor='#FFFFFF'><td class=inputname width=40%>
		<select name=col1>
		<% if col1<>"" then %>
			<option value="<%= col1 %>"><%= replace(replace(replace(replace(replace(col1,"**",""),"%%",""),"<>0","="),"1=1",""),"''","") %></option>
			<option value="">Select column to search on</option>
		<% else %>
			<option value="">Select column to search on</option>
		<% end if %>
		<option value="first_name like '%**%'">First Name like</option>
		<option value="last_name like '%**%'">Last Name like</option>
		<option value="sys_created > '**'">Registered After</option>
		<option value="sys_created < '**'">Registered Before</option>
		<option value="User_id like '%**%'">Login like</option>
		<option value="Company like '%**%'">Company like</option>
		<option value="Address like '%**%'">Address like</option>
		<option value="City like '%**%'">City like</option>
		<option value="State like '%**%'">State like</option>
		<option value="Zipcode like '%**%'">Zip Code like</option>
		<option value="Country like '%**%'">Country like</option>
		<option value="Phone like '%**%'">Phone like</option>
		<option value="Fax like '%**%'">Fax like</option>
		<option value="Email like '%**%'">Email like</option>
		<option value="Budget_left >= **">Budget >=</option>
		<option value="Budget_left <= **">Budget <=</option>
		<% if Service_Type>=7 then %>
		<option value="Reward_left >= **">Rewards >=</option>
		<option value="Reward_left <= **">Rewards <=</option>
		<% end if %>
		<option value="tax_id like '%**%'">Tax id like</option>
		<option value="spam<>0">Promo Emails =</option>
		<option value="is_residential<>0">Residential =</option>
		<option value="protected_page_access<>0">Protected Access =</option>
		<option value="ccid = **">Customer ID =</option>
		</td><td class=inputvalue><input type=text name=col1_value value="<%= col1_value %>" ID="Text1">
		<select name=col1_oper ID="Select3">
		<option value=AND>AND</option>
		<option value=OR>OR</option>
		</select></td></tr>
		<tr bgcolor='#FFFFFF'><td class=inputname width=40%>
		<select name=col2>
        <% if col2<>"" then %>
			<option value="<%= col2 %>"><%= replace(replace(replace(replace(replace(col2,"**",""),"%%",""),"<>0","="),"1=1",""),"''","") %></option>
			<option value="">Select column to search on</option>
        <% else %>
			<option value="">Select column to search on</option>
        <% end if %>
        <option value="first_name like '%**%'">First Name like</option>
        <option value="last_name like '%**%'">Last Name like</option>
        <option value="sys_created > '**'">Registered After</option>
        <option value="sys_created < '**'">Registered Before</option>
        <option value="User_id like '%**%'">Login like</option>
        <option value="Company like '%**%'">Company like</option>
        <option value="Address like '%**%'">Address like</option>
        <option value="City like '%**%'">City like</option>
        <option value="State like '%**%'">State like</option>
        <option value="Zipcode like '%**%'">Zip Code like</option>
        <option value="Country like '%**%'">Country like</option>
        <option value="Phone like '%**%'">Phone like</option>
        <option value="Fax like '%**%'">Fax like</option>
        <option value="Email like '%**%'">Email like</option>
        <option value="Budget_left >= **">Budget >=</option>
        <option value="Budget_left <= **">Budget <=</option>
        <% if Service_Type>=7 then %>
		<option value="Reward_left >= **">Rewards >=</option>
		<option value="Reward_left <= **">Rewards <=</option>
		<% end if %>
        <option value="tax_id like '%**%'">Tax id like</option>
        <option value="spam<>0">Promo Emails =</option>
        <option value="is_residential<>0">Residential =</option>
        <option value="protected_page_access<>0">Protected Access =</option>
        <option value="ccid = **">Customer ID =</option>

			</td><td class=inputvalue><input type=text name=col2_value value="<%= col2_value %>" ID="Text2"></td></tr>

			<tr align=center bgcolor='#FFFFFF'><td colspan="3">
				<input class="Buttons" name="SearchCustomers" type="submit" value="Search Customers">
				<% if Service_Type>=7 then %>
					<BR><input class="Buttons" name="Export_To_Text_File" type="submit" value="Export To Text" ID="Submit2">
					<input class="Buttons" name="Export_To_QuickBooks_File" type="submit" value="Export To Quickbooks" ID="Submit3">
				<% end if %></td></tr>
				<% if Request.Form("Export_To_QuickBooks_File") <> "" then %>
		<TR ><TD colspan="7" width="35%" align="center">Please read the <a href=quickbooks.asp class=link>Quickbooks instructions</a> and then Click on <b>Insert in QuickBook</b> button to Import data!!
                <BR><input class="Buttons" name="Insert In to Quickbooks" type="button"  value="Insert in Quick Book" onclick ="Display()" ID="Button1"></td></tr>
<% end if %>
			
		</table></TD></TR>

<% if (request.querystring<>"" or request.form<>"") AND Request.Form("Export_To_QuickBooks_File")= "" then %>
<!--#include file="list_view.asp"-->
<%
else
        if myStructure("AddAllowed") then %>
	       <tr bgcolor="#FFFFFF"><td><input type="button" OnClick=JavaScript:self.location="<%= myStructure("NewRecord") %>" class="Buttons" value="Add Customer" name="Create_new"></td></tr>
	<% end if
end if

Delete_Id = fn_get_delete_ids
if Delete_Id<>"" then

	sql_check ="select oid From Store_Purchases WITH (NOLOCK) Where cid in ("&Delete_Id&") and store_id = "&store_id
  	set myfields=server.createobject("scripting.dictionary")
  	Call DataGetrows(conn_store,sql_check,mydata,myfields,noRecords)

	sOid=""
  	if noRecords = 0 then
  		FOR rowcounter= 0 TO myfields("rowcount")
			if rowcounter=0 then
				sOid=mydata(myfields("oid"),rowcounter)
			else
				sOid=sOid&","&mydata(myfields("oid"),rowcounter)
			end if
  		Next
		sql_delete = "delete from store_purchases where store_Id="&Store_Id&" and Oid in ("&sOid&");"&_
			"delete from store_transactions where store_Id="&Store_Id&" and Oid in ("&sOid&");"&_
			"delete from store_purchases_returns where store_Id="&Store_Id&" and Oid in ("&sOid&");"&_
			"delete from store_gift_certificates_dets where store_Id="&Store_Id&" and Order_id in ("&sOid&");"&_
                        "delete from store_purchases_shippments where store_Id="&Store_Id&" and Oid in ("&sOid&")"
			
		conn_store.Execute sql_delete
  	End if
end if
	
createFoot thisRedirect, 2

%>


