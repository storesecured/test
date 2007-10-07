<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the payment mode as check for the reseller site.
'	Page Name:		   	reseller_payment_mode.asp
'	Version Information:EasystoreCreator
'	Input Page:		    resellersignup.asp
'	Output Page:	    reseller_payment_mode.asp
'	Date & Time:		31 Aug 1 Sept 2004	
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 


%>

<%

if trim(Request.Form("hidsession"))="" and  trim(session("ResellerID")) = "" then 
	Response.redirect "reseller_error.asp?error=4"
end if


title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="reseller_payment_mode"
includejs=1
'REPLACE 'SINGLE QUOTES
function fixquotes(str)
fixquotes = replace(str&"","'","''")
end function

%>
<!--#include file="header.asp"-->

<script language="javascript">
function fnChange()
{
		var val
		val = document.frmaddress.modeselect.value
				
		document.frmaddress.action="reseller_payment_mode.asp?action=show&val="+val;
		document.frmaddress.submit();
	
}

function fnCreate()
{
	var ErrMsg
	ErrMsg=""
	
	if(document.frmaddress.modeselect.value==1)
	//Validations for blank & valid email id.
	{
		if (isWhitespace(document.frmaddress.txtEmail.value)== true)
		{
			ErrMsg = ErrMsg +"Email field is mandatory.\n";
		}
		if (isWhitespace(document.frmaddress.txtEmail.value)== false)
		{
			if(IsEmail(document.frmaddress.txtEmail.value)==false)
			{
				ErrMsg = ErrMsg + "Invalid Email id. \n" ;		
			}	
		}

	}
	
	
	
	if(document.frmaddress.modeselect.value==2)
	{
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
		if (isWhitespace(document.frmaddress.txtAddress.value)== false)
		{
			if(isAlphaNumeric(document.frmaddress.txtAddress.value)==false)
			{
				ErrMsg = ErrMsg + "Address should not contain special characters.\n" ;		
			}	
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
		//		ErrMsg = ErrMsg + "ZipCode should be numeric. \n" ;		
			}
		}
	
		
		
		if (document.frmaddress.selcountry.value == "0")
		{
	
			ErrMsg = ErrMsg + "Country is manadatory. \n" ;		
		}
	
		
	}
		

	
		if (ErrMsg!="")
		{
			alert(ErrMsg);
		}
		else
		{
			document.frmaddress.action="reseller_payment_mode.asp?action=save";
			document.frmaddress.submit();
		}
}

</script>
<%
'To retrieve reseller id from session variable
if trim(Request.Form("hidsession"))<> "" then 
	session("ResellerID") = trim(Request.Form("hidsession"))
	intresellerid = session("ResellerID") 
end if
if session("ResellerID")<> "" then 
	intresellerid = session("ResellerID")
end if	

'code here to show the fields according to the mode select
if trim(Request.QueryString("action"))="show" then
	mode=trim(Request.QueryString("val"))
end if


if trim(Request.QueryString("action"))="save" then
	'Retriving the values from fields
	mode=trim(Request.Form("modeselect"))
	if mode="1" then 
			email = trim(Request.Form("txtEmail"))
	end if	
	if mode="2" then 
		payble = trim(fixquotes(Request.Form("txtName")))
		address = trim(fixquotes(Request.Form("txtAddress")))
		state = trim(fixquotes(Request.Form("txtState")))
		city = trim(fixquotes(Request.Form("txtCity")))
		zip = trim(fixquotes(Request.Form("txtZipCode")))
		country = trim(Request.Form("selcountry"))
	end if
	if mode = "1" then 'if some mode is selected the insert the values into the payment mode table
		sqlnew="insert into TBL_Esc_Reseller_payment_mode(fld_reseller_id,fld_payment_mode,fld_email_address_for_paypal)"&_
			  " values("&intResellerID&",1,'"&email&"')"
		conn.execute (sqlnew)
	end if
	
	if mode = "2" then 'if some mode is selected the insert the values into the payment mode table
		sqlnew  = "insert into TBL_Esc_Reseller_payment_mode(fld_reseller_id,fld_payble_to,fld_address,fld_state,fld_city,fld_zipcode,fld_country,"&_
				  " fld_payment_mode) values ("&intResellerID&",'"&payble&"','"&address&"','"&state&"','"&city&"','"&zip&"','"&country&"',2)"
		conn.execute (sqlnew)
	end if
	
	
	
	%>
	
	<script language="">
	document.location.href="intermediate.asp"
	</script>
	<%
		
end if
'To retrive country name & country id from the database
strcountry="select Country,Country_id from sys_countries where country <> 'All Countries' ORDER BY Country"
set rscountry=conn.execute(strcountry)
%>

	
<table border="0" width=525 align=left bordercolor="#000066">

	<tr>
		<td colspan="2" valign="top" align=left>
		<b>Reseller Choose Payment Mode</b><br><br>
		Please indicate how you would like to receive payment for your active clients. <br><br>
		</td>
	</tr>

           <FORM action="" method=post name="frmaddress">
              <tr>
              <td>
              <b>Choose mode </b>
              </td>
              <td>
              <select name="modeselect" onchange="javacript:fnChange()" >
              <option value="0">--Select Mode--</option>
              <option value="1" <%if mode="1" then%>selected<%end if%>>By Paypal</option>
              <option value="2" <%if mode="2" then%>selected<%end if%>>By Check</option>
              </select>
              </td>
              </tr>
              <tr>
              	<tr>
				<%if mode="1" and trim(Request.QueryString("action"))="show" then %>
				<td>
					  Email address
				</td>
				<td>
					 <input type="text" name="txtEmail" maxlength="100" >
				</td>	
			  </tr>	
				<%end if%>
			  <%if mode="2" and trim(Request.QueryString("action"))="show" then %>
			  <tr>
        		<td>
					  Payable to 
				</td>
				<td>
					 <input type="text" name="txtName" maxlength="50" >
				</td>	
				</tr>	
				<tr>
            	 <td>
					  Address
				</td>
				<td>
					 <input type="textarea" name="txtAddress" maxlength="100" >
				</td>	
				</tr>	
				<tr>
            	 <td>
					  State
				</td>
				<td>
					 <input type="text" name="txtState" maxlength="50" >
				</td>	
				</tr>	
				<tr>
            	 <td>
					  City
					  </td>
					  <td>
					 <input type="text" name="txtCity" maxlength="100" >
				</td>	
				</tr>	
					
			<tr>
			
            	 <td>
					  ZipCode
					  </td>
					  <td>
					 <input type="text" name="txtZipCode" maxlength="20" >
				</td>	
				</tr>	
				
			<tr>
			<td>
			Country
			</td>
			<td>
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
			<% end if%>
			<%' if trim(Request.QueryString("action"))="show" then%>
			<tr>
			<td>
				<input type="button" name="save" value="Save" onclick="javascript:fnCreate()">
				<input type="Reset" name="Reset" value="Reset">
			</td>
			</tr>
			<%' end if%>
            	 
			</form>
	</table>
<!--#include file="footer.asp"-->
