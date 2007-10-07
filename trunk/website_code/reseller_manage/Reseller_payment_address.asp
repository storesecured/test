<!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the payment mode as check for the reseller site.
'	Page Name:		   	reseller_payment_address.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_payment_address.asp
'	Date & Time:		5 Aug & 16 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 

mSel7=3
mSel8=2
%>


<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="../include/commonfunctions.js"></SCRIPT>
<script language="javascript">
function fnsave()
{
	var ErrMsg
	ErrMsg=""
	
	//Validations for blank & valid email id.
		if (isWhitespace(document.frmaddress.txtName.value)== true)
		{
			ErrMsg = ErrMsg +"Payable to is mandatory.\n";
		}
		if (isWhitespace(document.frmaddress.txtName.value)== false)
		{
			if(	isAlphaNumeric(document.frmaddress.txtName.value)==false)
			{
				ErrMsg = ErrMsg + "Payable to should not contain special characters.\n" ;		
			}	
		}
		if (isWhitespace(document.frmaddress.txtAddress.value)== true)
		{
			ErrMsg = ErrMsg +"Address is mandatory.\n";
		}
		
		if (isWhitespace(document.frmaddress.txtState.value)== true)
		{
			ErrMsg = ErrMsg +"State is mandatory.\n";
		}
		
		if (isWhitespace(document.frmaddress.txtState.value)== false)
		{
			if(isAlphaNumeric(document.frmaddress.txtState.value)==false)
			{
				ErrMsg = ErrMsg + "State should not contain special characters.\n" ;		
			}	
		}
	
		if (isWhitespace(document.frmaddress.txtCity.value)== true)
		{
			ErrMsg = ErrMsg +"City is mandatory.\n";
		}
		
		
		if (isWhitespace(document.frmaddress.txtCity.value)== false)
		{
			if(isAlphaNumeric(document.frmaddress.txtCity.value)== false)
			{
				ErrMsg = ErrMsg + "City  should not contain special characters.\n" ;		
			}	
		}
		
		if (isWhitespace(document.frmaddress.txtZipCode.value )== true)
		{
			ErrMsg = ErrMsg +"ZipCode is mandatory.\n";
		}
		
		
		if (isWhitespace(document.frmaddress.txtZipCode.value) == false)
		{
			if (isAllNumeric(document.frmaddress.txtZipCode.value) == false)
			{
			//	ErrMsg = ErrMsg + "ZipCode should be numeric. \n" ;		
			}
		}
		if (document.frmaddress.selcountry.value== 0)
		{
			ErrMsg = ErrMsg +"Country is mandatory.\n";
		}
		
		
		if (ErrMsg!="")
		{
			alert(ErrMsg);
		}
		else
		{
			document.frmaddress.action="Reseller_payment_address.asp?action=save";
			document.frmaddress.submit();
		}
}

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
dim strinsert,intResellerID,stremail,intFlag
intFlag =0
'retreving the reseller id from the session
intResellerID = trim(session("ResellerID"))


'check here if the reseller has already made any entry for check(hardcode check for payment mode=2"
'-------------------------------------------------------------------------------
stremail = " select fld_payble_to,fld_address,fld_city,fld_state,fld_country,fld_zipcode from tbl_esc_reseller_payment_mode where fld_reseller_id="&intResellerID&" and "&_
		   " fld_payment_mode=2"
set rsemail=conn.execute(stremail)
if not rsemail.eof then
		payble=trim(rsemail(checkencode("fld_payble_to")))
		address=trim(rsemail(checkencode("fld_address")))
		city=trim(rsemail(checkencode("fld_city")))
		state=trim(rsemail(checkencode("fld_state")))
		country=trim(rsemail("fld_country"))
		zip=trim(checkencode(rsemail("fld_zipcode")))
end if
rsemail.close	
set rsemail = Nothing

'To retrive country name & country id from the database
'strcountry="select fld_Country_name,fld_Country_id from tbl_country order by fld_Country_id asc"
strcountry="select Country,Country_id from sys_countries where country <> 'All Countries' ORDER BY Country"
set rscountry=conn.execute(strcountry)

if trim(request.querystring("action"))="save"  then 
	'check here ifxx the reseller has already made any entry paypal or for check(irrespective of payment mode)
	'-------------------------------------------------------------------------------
	
	stremail = " select fld_payble_to,fld_address,fld_city,fld_state,fld_country,fld_zipcode from tbl_esc_reseller_payment_mode where fld_reseller_id="&intResellerID&""
	
	'Retriving the values of all fields
			strpay=trim(Request.Form("txtName"))
			strpay=fixquotes(strpay)
			stradd=trim(Request.Form("txtAddress"))
			stradd=fixquotes(stradd)
			strstate=trim(Request.Form("txtState"))
			strstate=fixquotes(strstate)
			strcity=trim(Request.Form("txtCity"))
			strcity=fixquotes(strcity)
			strzip=trim(fixquotes(Request.Form("txtZipCode")))
			strcoun=trim(Request.Form("selCountry"))
			
	set rscheck = conn.execute (stremail)  
	if not rscheck.eof then
		
		
		'Entry is there

			'Update query for that reseller
			
		strupdate= "Update TBL_Esc_Reseller_payment_mode set fld_payble_to='"&strpay&"',fld_address='"&stradd&"',"&_
				   " fld_city='"&strcity&"',fld_state='"&strstate&"',fld_country='"&strcoun&"',fld_zipcode='"&strzip&"', "&_
				   " fld_payment_mode=2 where fld_reseller_id="&intResellerID&""
		conn.execute (strupdate)
		intFlag = "1"
		else
		
		'Entry is not there
		' insert query for that reseller
				'Inserting values into the table
			strinsert = "insert into TBL_Esc_Reseller_payment_mode(fld_reseller_id,fld_payble_to,fld_address,"&_
					    " fld_state,fld_city,fld_zipcode,fld_country,fld_payment_mode) values ("&intResellerID&",'"&strpay&"','"&stradd&"','"&strstate&"','"&strcity&"','"&strzip&"','"&strcoun&"',2)"
			conn.execute (strinsert)
		intFlag = "2"
		end if
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
          <TD class=title><B>Reseller</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%">
            <UL><BR><BR></UL></TD></TR></TBODY></TABLE></TD></TR>
            <tr>
            <td>
      <!--#include file="incmenu.asp"-->
            </td>
            </tr>
            
            
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
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="Reseller_sales_history.asp">View Sales History
              </A></TD>
        </TR>
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="Reseller_payment_Mode.asp">Choose Payment Mode
              </A></TD>
        </TR>   
        
          <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="edit_reseller.asp">Edit Profile
              </A></TD>
        </TR>   
     
            
                          
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Payment Mode - Check
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action="" method=post name="frmaddress">
              <tr>
               <br>
				 <td>
					 <b> Enter the information for sending check </b>
				</td>
				<td>
					 
				</td>	
				</tr>	
				<tr>
            <tr>
             <br>
				 <td>
					  Payable to 
				</td>
				<td>
					 <input type="text" name="txtName" maxlength="100" value="<%= payble%>">
				</td>	
				</tr>	
				<tr>
            	 <td>
					  Address
				</td>
				<td>
					 <input type="textarea" name="txtAddress" maxlength="100" value="<%= address%>">
				</td>	
				</tr>	
				<tr>
            	 <td>
					  State
				</td>
				<td>
					 <input type="text" name="txtState" maxlength="50" value="<%= state%>">
				</td>	
				</tr>	
				<tr>
            	 <td>
					  City
				</td>
				<td>
					 <input type="text" name="txtCity" maxlength="50" value="<%= city%>">
				</td>	
				</tr>	
					
			<tr>
			
            	 <td>
					  ZipCode
				</td>
				<td>
					 <input type="text" name="txtZipCode" maxlength="20" value="<%= zip%>">
				</td>	
				</tr>	
<tr>
			
            	 <td>
					  Country
				</td>
				<td>
					<!-- <input type="text" name="txtCountry" maxlength="100" value="<%= country%>">-->
					<%
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
					
				<tr>
				<td>
					  
				</td>
				<td>
				<br>
				           	
					 <INPUT name="save" type="button" value="Save" onclick="javascript:fnsave()"> 
					 <input type="reset" type="reset" value="Reset">
					 <input type="button" value="Back" name="back" onclick="javascript:history.back()">
				</td>	
              
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
<% if intFlag="1" then%>
<script language=javascript>
	alert("Resellers records have been updated.")
	document.location.href = "reseller_payment_mode.asp"
</script>
<% end if %>

<% if intFlag="2" then%>
<script language=javascript>
	alert("Resellers records have been added.")
	document.location.href = "reseller_payment_mode.asp"
</script>
<% end if %>