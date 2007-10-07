<!--#include file = "include/escheader.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page will edit the information about the reseller.
'	Page Name:		   	edit_reseller.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_list.asp
'	Output Page:	    edit_reseller.asp
'	Date & Time:		10 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	##############
mSel3=3
mSel4=2
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="../include/commonfunctions.js" 
type=text/javascript>
</SCRIPT>
<script language=javascript>
function fnsave()
{
	var ErrMsg
	ErrMsg=""

	//Validations for blank fields
	if (isWhitespace(document.frmedit.txtcompany.value)== true)
		{
			//ErrMsg = ErrMsg + "Company name is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtcompany.value)== false)
		{
			if(	isAlphaNumeric(document.frmedit.txtcompany.value)==false)
			{
				ErrMsg = ErrMsg + "Company should not contain special characters.\n" ;		
			}	
		}
		if (isWhitespace(document.frmedit.txtFirstName.value)== true)
		{
			ErrMsg = ErrMsg + "First name is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtFirstName.value)== false)
		{
			if(	isAlphaNumeric(document.frmedit.txtFirstName.value)==false)
			{
				ErrMsg = ErrMsg + "First name should not contain special characters.\n" ;		
			}	
		}
		
		if (isWhitespace(document.frmedit.txtLastName.value)== true)
		{
			ErrMsg = ErrMsg + "Last name is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtLastName.value)== false)
		{
			if(	isAlphaNumeric(document.frmedit.txtLastName.value)==false)
			{
				ErrMsg = ErrMsg + "Last name should not contain special characters.\n" ;		
			}	
		}
		
		if (isWhitespace(document.frmedit.txtAddress.value)== true)
		{
			ErrMsg = ErrMsg + "Address is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtCity.value)== true)
		{
			ErrMsg = ErrMsg + "City is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtCity.value)== false)
		{
			if(isAllCharacters(document.frmedit.txtCity.value)==false)
			{
				ErrMsg = ErrMsg + "City should contain characters only.\n" ;		
			}	
		}
		if (isWhitespace(document.frmedit.txtEmail.value)== true)
		{
			ErrMsg = ErrMsg + "Email is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtEmail.value)== false)
		{
			if(IsEmail(document.frmedit.txtEmail.value)==false)
			{
				ErrMsg = ErrMsg + "Invalid Email id. \n" ;		
			}	
		}
		/*
		if (isWhitespace(document.frmedit.txtCardNumber.value)== true)
		{
			ErrMsg = ErrMsg + "Credit Card Number is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmedit.txtCode.value)== true)
		{
			ErrMsg = ErrMsg + "Card Code is mandatory. \n" ;		
		}
		*/
		if (isPhone(document.frmedit.txtFax.value)== false)
		{
			ErrMsg = ErrMsg + "Fax should be numeric. \n" ;		
		}
		
		if(isPhone(document.frmedit.txtPhone.value)== false)
		{
			ErrMsg = ErrMsg + "Phone should be numeric & valid. \n" ;		
		}
		
		if (isAllNumeric(document.frmedit.txtZipCode.value)== false)
		{
			//ErrMsg = ErrMsg + "ZipCode should be numeric. \n" ;		
		}
		/*
		if(document.frmedit.mm.value==0)
		{
			ErrMsg = ErrMsg + "Month is mandatory.\n"		
		}
	
		if (document.frmedit.yy.value==0)
		{
			ErrMsg = ErrMsg +"Year is mandatory.\n "		
		}*/
	
		if (ErrMsg!="")
		{
		alert(ErrMsg);
		}
		
		else
		{
			document.frmedit.action="edit_reseller.asp?saveaction=save";
			document.frmedit.submit();
		}
}
</script>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%

	
dim intresellerid,sqlGetData,strcompany,strfirst,strlast,straddress,strcity,strstate,sqlupdate,company,first,last,address,city,state,zip,phone,fax,emailcardno,intmonth,intyear,code,intflag

			'Retriving the values from the text boxes
			
if trim(request.querystring("saveaction"))="save"  then 
	company = trim(Request.Form("txtcompany"))
	first=trim(Request.Form("txtFirstName"))
	last=trim(Request.Form("txtLastName"))
	address=trim(Request.Form("txtAddress"))
	city = trim(Request.Form("txtCity"))
	state=trim(Request.Form("txtState"))
	company  = fixquotes(company)
	first  = fixquotes(first)
	last  = fixquotes(last)
	address  = fixquotes(address)
	city  = fixquotes(city)
	state  = fixquotes(state)
	zip=trim(fixquotes(Request.Form("txtZipCode")))
	phone=trim(Request.Form("txtPhone"))
	fax=trim(Request.Form("txtFax"))
	email = trim(Request.Form("txtEmail"))
	cardno=trim(Request.Form("txtCardNumber"))
	intmonth=  trim(Request.Form("mm"))
	intyear=trim(Request.Form("yy"))
	code=trim(Request.Form("txtCode"))
	country=trim(Request.Form("selcountry"))
	intresellerid = trim(Request.Form("HidResellerID"))
	
	'Updating the records of the reseller
	if not rsGetData.eof then
		sqlupdate = "update tbl_reseller_master set fld_company_name='"&company&"',fld_first_name='"&first&"',fld_last_name='"&last&"'"&_
					",fld_address='"&address&"',fld_city='"&city&"',fld_state='"&state&"',fld_zip_code='"&zip&"',fld_phone='"&phone&"',fld_fax='"&fax&"'"&_
					",fld_mail='"&email&"',fld_country='"&country&"' where fld_reseller_id="&intresellerid&""
					conn.execute(sqlupdate)
	
		
		intflag=1
	end if
end if

	if trim(Request.QueryString("action"))<>"" then
	intresellerid = trim(Request.QueryString("action"))
		sqlGetData = "select fld_company_name,fld_first_name,fld_last_name,fld_address,fld_city,fld_state,fld_zip_code,"&_
					 " fld_phone,fld_fax,fld_mail,fld_country from tbl_reseller_master where fld_reseller_id="&intresellerid&""
					 
		set rsGetData = conn.execute(sqlGetData)
		'Retriving data about reseller from database
		
			if not rsGetData.eof then
			strcompany = trim(checkencode(rsGetData("fld_company_name")))
			strfirst = trim(checkencode(rsGetData("fld_first_name")))
			strlast = trim(checkencode(rsGetData("fld_last_name")))
			straddress = trim(checkencode(rsGetData("fld_address")))
			strcity = trim(checkencode(rsGetData("fld_city")))
			strstate = trim(checkencode(rsGetData("fld_state")))
			strzip = trim(checkencode(rsGetData("fld_zip_code")))
			strphone = trim(rsGetData("fld_phone"))
			strfax = trim(rsGetData("fld_fax"))
			strmail = trim(rsGetData("fld_mail"))
			strcardno = trim(rsGetData("fld_credit_card_number"))
			strExpireMonth = trim(rsGetData("fld_card_expire_month"))
			strExpireYear = trim(rsGetData("fld_card_expire_year"))
			strcardcode = trim(rsGetData("fld_card_code"))
			country = trim(rsGetData("fld_country"))
			end if
			rsGetData.close
			rsGetData=nothing

	end if
%>
<DIV id=overDiv 
style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  <TBODY>
  <TR>
    <TD class=title>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
          <TD class=title><B>Easystorecreator</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%">
            <UL><BR><BR></UL></TD></TR></TBODY></TABLE></TD></TR>
          

  <TR>
    <TD><!--#include file="incESCmenu.asp"--></TD></TR>
  <TR>
    <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        </TBODY></TABLE></TD></TR>
  <TR>
    <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR vAlign=top>
          <TD rowSpan=2 width=180>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>          
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <FORM method=post name="frmedit">
          <table border="1">
          <tr>
          Resellers Personal Info 
		<hr noshade>	
          </tr>
          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Company Name</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name="txtcompany" value="<%= strcompany%>" > 
          </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>First Name</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtFirstName value="<%= strfirst%>" > 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Last Name</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtLastName value="<%= strlast%>" > 
            </TD></TR>
        
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Address</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtAddress value="<%= straddress%>" > 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>City</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtCity value="<%= strcity%>"> 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>State</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtState value="<%= strstate%>"> 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>ZipCode</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtZipCode value="<%= strzip%>" > 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Phone</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtPhone value="<%= strphone%>" > 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Fax</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtFax value="<%= strfax%>"> 
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Email</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtEmail value="<%= strmail%>" > 
            </TD></TR>
        <tr>
        <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Country</FONT></TD>
          <TD width="1%"></TD>
        <td>
        <%
        'To retrive country name & country id from the database
		strcountry="select Country,Country_id from sys_countries where country <> 'All Countries' ORDER BY Country"
		set rscountry=conn.execute(strcountry)
        if not rscountry.eof then%>
					<select name="selcountry">
					<option  value="0" <%=sel%>>--Select Country--</option>
					<% while not rscountry.eof
				
					countryname=trim(rscountry("Country"))
					
					countryid=trim(rscountry("Country_id"))
					
						if country=trim(rscountry("Country_id")) then 
							 sel ="selected"
						else
							sel = ""
						end if
					
						%>
						<option  value="<%= countryid%>" <%=sel%>><%= countryname%></option>
					
						<%
					
						rscountry.movenext
							wend
					end if 
					rscountry.close
					set rscountry=nothing
					%>
					</select>
		</td>
		</tr>			
        <!--<TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Credit Card Number</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtCardNumber value="<%= strcardno%>" > 
            </TD></TR>
             <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Expiration Date</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <select name="mm" size="1">
          
				   <option value=0>Select month</option>
				   <option value="01" <%if strExpireMonth="1" then %>selected<%end if%>>January</option>
				   <option value="02" <%if strExpireMonth="2" then %>selected<%end if%>>February</option>
				   <option value="03" <%if strExpireMonth="3" then %>selected<%end if%>>March</option>
				   <option value="04" <%if strExpireMonth="4" then %>selected<%end if%>>April</option>
				   <option value="05" <%if strExpireMonth="5" then %>selected<%end if%>>May</option>
				   <option value="06" <%if strExpireMonth="6" then %>selected<%end if%>>June</option>
				   <option value="07" <%if strExpireMonth="7" then %>selected<%end if%>>July</option>
				   <option value="08" <%if strExpireMonth="8" then %>selected<%end if%>>August</option>
				   <option value="09" <%if strExpireMonth="9" then %>selected<%end if%>>September</option>
				   <option value="10" <%if strExpireMonth="10" then %>selected<%end if%>>October</option>
				   <option value="11" <%if strExpireMonth="11" then %>selected<%end if%>>November</option>
				   <option value="12" <%if strExpireMonth="12" then %>selected<%end if%>>December</option>
				   </select>
				   <% Current_year = year(now()) %>
				   <select name="yy" size="1">
				   <option value="0">Select year</option>
				   <% sel=""
					if strExpireYear< 10 then 
						strExpireYear = 0&strExpireYear
				   end if
			  
				   for go_year = Current_year to Current_year + 10 %>
				         <% go_year_two_digit = right(go_year,2) 
				        
				        if trim(go_year_two_digit)=trim(strExpireYear) then
							sel="selected"
						else
							sel = ""
						end if	
				         
				         %>
				         
				         <%=sel%><option value="<%= go_year_two_digit %>" <%=sel%>><%= go_year %></option>
				      <% next %>
				</select>
				</td>

            </TR>
            
           

            <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Card Code</FONT> </TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=200 name=txtCode value="<%= strcardcode%>"> 
            </TD></TR>-->
            <input type="hidden" value="<%=intresellerid%>" name="hidResellerID">
        <tr>
        <TD noWrap width="1%">&nbsp;</TD>
        <td noWrap width="1%">
		<INPUT name="save" type="button" value="Save" onclick="javascript:fnsave()">&nbsp; 
        <input type="reset" value="Reset" name="Reset"> 
        <input type="button" value="Back" name="back" onclick="javascript:history.back()">
				</td></tr>	
		  
          </table>
          
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              
              <TBODY>
              <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
<% if intflag="1" then%>
<script language=javascript>
	alert("Resellers records have been updated.")
	document.location.href = "reseller_list.asp"
</script>
<% end if %>
