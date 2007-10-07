

<%

 

if trim(Session("tempresellerid")) ="" then 
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
tracking_page_name="reseller_billing_info"
includejs=1
%>
<!--#include file="header.asp"-->

<%
dim strinsert
'Retriving card type selected from reseller_payment_page
cardtype = trim(Request.Form("payment_method"))
'Retriving reseller id from session variable
resellerid = trim(Request.Form("hidSessionID"))

if trim(Request.QueryString("infoaction"))<> "save" then
	'Updating card type for that reseller
	str="update tblTemp set fld_card_type='"&cardtype&"' where fld_reseller_id="&resellerid&""
	conn.Execute str
end if
'To retrive country name & country id from the database
strcountry="select Country,Country_id from sys_countries where country <> 'All Countries' ORDER BY Country"
set rscountry=conn.execute(strcountry)

%>

<script language="javascript" src="include/commonfunctions.js"></script>
<script language=javascript >

		

function fnshow(val)
{
	var ErrMsg
	ErrMsg = ""
	val = val
	if(isWhitespace(document.frmbilling.First_name.value)==true)
	{
		ErrMsg = ErrMsg + "First name is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.First_name.value)==false)
	{
		if(isAlphaNumeric(document.frmbilling.First_name.value)==false)
		{
				ErrMsg = ErrMsg + "First name should not contain special characters.\n" ;		
		}
	}
	
	if(isWhitespace(document.frmbilling.Last_name.value)==true)
	{
		ErrMsg = ErrMsg + "Last name is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.Last_name.value)==false)
	{
		if(isAlphaNumeric(document.frmbilling.Last_name.value)==false)
		{
				ErrMsg = ErrMsg + "Last name should not contain special characters.\n" ;		
		}
	}
	
	if(isWhitespace(document.frmbilling.Address.value)==true)
	{
		ErrMsg = ErrMsg + "Address is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.City.value)==true)
	{
		ErrMsg = ErrMsg + "City is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.City.value)==false)
	{
		if(isAlphaNumeric(document.frmbilling.City.value)==false)
		{
				ErrMsg = ErrMsg + "City should not contain special characters.\n" ;		
		}
	}
	
	if(isWhitespace(document.frmbilling.State.value)==true)
	{
		ErrMsg = ErrMsg + "State is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.State.value)==false)
	{
		if(isAlphaNumeric(document.frmbilling.State.value)==false)
		{
				ErrMsg = ErrMsg + "State should not contain special characters.\n" ;		
		}
	}
	
	if(isWhitespace(document.frmbilling.Zip.value)==true)
	{
		ErrMsg = ErrMsg + "Zip is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.Zip.value)==false)
	{
		if(isPhone(document.frmbilling.Zip.value)==false)
		{
	//		ErrMsg = ErrMsg + "Zip should be valid.\n";
		}
	}
	
	if(document.frmbilling.selcountry.value==0)
	{
		ErrMsg = ErrMsg + "Country is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.Phone.value)==true)
	{
		ErrMsg = ErrMsg + "Phone is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.Phone.value)==false)
	{
		if(isPhone(document.frmbilling.Phone.value)==false)
		{
			ErrMsg = ErrMsg + "Phone should be valid.\n";
		}
	}   
	
	if(isWhitespace(document.frmbilling.Fax.value)==true)
	{
	//	ErrMsg = ErrMsg + "Fax is mandatory.\n";
	}
	
	if(isWhitespace(document.frmbilling.Fax.value)==false)
	{
		if(isPhone(document.frmbilling.Fax.value)==false)
		{
			ErrMsg = ErrMsg + "Fax should be valid.\n";
		}
	}	
	if(isWhitespace(document.frmbilling.Company.value)==true)
	{
		//ErrMsg = ErrMsg + "Company is mandatory.\n";
	}
	
	
	if (val==1)
	{
		if(isWhitespace(document.frmbilling.cardno.value)==true)
		{
			ErrMsg = ErrMsg + "Credit Card Number is mandatory.\n";
		}
	
		if(isWhitespace(document.frmbilling.cardcode.value)==true)
		{
			//ErrMsg = ErrMsg + "Card Code is mandatory.\n";
		}
	
		if(document.frmbilling.month.value==0)
		{
			ErrMsg = ErrMsg + "Month is mandatory.\n";
		}
	
		if(document.frmbilling.year.value==0)
		{
			ErrMsg = ErrMsg + "Year is mandatory.\n";
		}
	}
	
	if (val==2)
	{	
		if(isWhitespace(document.frmbilling.bank.value)==true)
		{
			ErrMsg = ErrMsg + "Bank Name is mandatory.\n";
		}
	
		if(isWhitespace(document.frmbilling.routing.value)==true)
		{
			ErrMsg = ErrMsg + "Routing is mandatory.\n";
		}

		if(isWhitespace(document.frmbilling.account.value)==true)
		{
			ErrMsg = ErrMsg + "Account is mandatory.\n";
		}

		if(isWhitespace(document.frmbilling.check.value)==true)
		{
			ErrMsg = ErrMsg + "Check is mandatory.\n";
		}
	
		if(isWhitespace(document.frmbilling.lic.value)==true)
		{
			ErrMsg = ErrMsg + "Drivers Licence is mandatory.\n";
		}
	}
	
	if (val==3)
	{
		if(isWhitespace(document.frmbilling.email.value)==true)
		{
			ErrMsg = ErrMsg + "Email id is mandatory.\n";
		}
		
		if(isWhitespace(document.frmbilling.email.value)==false)
		{
			if(IsEmail(document.frmbilling.email.value)==false)
			{
					ErrMsg = ErrMsg + "Invalid Email id. \n" ;	
			}
		}
	}

	if(ErrMsg!="")
	{
		alert(ErrMsg);
	}
	
	else
	{
		
		document.frmbilling.action="reseller_billing_info2.asp";
		//document.frmbilling.action="../manage/reseller_billing_info.asp?infoaction=save";
		document.frmbilling.submit();
	}
}
</script>

<%if cardtype <> "Paypal" then%>

<table border="0" width=525 align=left bordercolor="#000066">

			<tr>
			<td valign="top" align=center>
					<br>
				</tr>
		<b>Reseller Program</b>
			<form method="POST" action="" name="frmbilling"> 
			<tr>
				<td>First Name</td>
				<td><input name="First_name" maxlength=50></td>
			</tr>
			<tr>
				<td>Last Name</td>
				<td><input name="Last_name" maxlength=50></td>
			</tr>
			
			
			<tr>
				<td>Address</td>
				<td><input name="Address" maxlength=50></td>
			</tr>
			<tr>
				<td>City</td>
				<td><input name="City" maxlength=50></td>
			</tr>
			<tr>
				<td>State</td>
				<td><input name="State" maxlength=50></td>
			</tr>
			
			<tr>
				<td>Zip</td>
				<td><input name="Zip" maxlength=50></td>
			</tr>
			
			<tr>
				<td>Phone</td>
				<td><input name="Phone" maxlength=50></td>
			</tr>
			<tr>
				<td>Fax</td>
				<td><input name="Fax" maxlength=50></td>
			</tr>
				<tr>
				<td>Company</td>
				<td><input name="Company" maxlength=50></td>
			</tr>
		
			<tr>
				<td>Country</td>
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
		
			
<%end if%>			
				<% if cardtype = "Visa" or cardtype = "Mastercard" or cardtype =  "American Express" or cardtype = "Discover" then 
				intflag = "1"
				%>
			<tr>
			<td>Credit Card Number</td>
			<td><input type="text" value="" name="cardno">&nbsp;
			
			<%if cardtype = "Visa" then%>
			<IMG src="images/icon_visa.gif">
			<%end if%>
			
			<% if cardtype = "Mastercard" then%>
				<IMG src="images/icon_mastercard.gif">
			<%end if%>
			
			<% if cardtype = "American Express" then%>
			<IMG src="images/icon_amex.gif">
            <%end if%>     
            
			<% if cardtype = "Discover" then%>
			<IMG src="images/icon_discover.gif">
			<% end if%>
			
			<TD class=inputvalue>
			</td>
			</tr>
			<tr>
			<TD class="inputname">Expiration Date</TD>
			<td>
			<select name="month" size="1">
				<option selected value="0" >Select Month</option>
				<option value="01">January</option>
				<option value="02">February</option>
				<option value="03">March</option>
				<option value="04">April</option>
				<option value="05">May</option>
				<option value="06">June</option>
				<option value="07">July</option>
				<option value="08">August</option>
				<option value="09">September</option>
				<option value="10">October</option>
				<option value="11">November</option>
				<option value="12">December</option>
			</select>&nbsp;
			 <% Current_year = year(now()) %>
    		<select name="year" size="1">
			   <option selected value="0">Select Year</option>
			   <% for go_year =  Current_year to Current_year + 10 %>
			         <% go_year_two_digit = right(go_year,2) %>
			         <option value="<%= go_year_two_digit %>"><%= go_year %>
			   <% next %>
			</select>
			</td>
			<td>
       
			</td>
			</tr>
			
			<tr>
			<td>
			Card Code
			</td>
			<td>
			<input type="text" value="" name="cardcode" size="4" maxlength="4">
			</td>
			</tr>
			
			<%end if%>
			
			<% if cardtype = "eCheck" then
			intflag = "2"
			%>
			<tr>
			<td>
			Bank Name
			</td>
			<td>
			<input type="text" value="" name="bank">
			</td>
			<tr>
			<td>
			Routing #
			</td>
			<td>
			<input type="text" value="" name="routing">
			</td>
			</tr>
			
			<tr>
			<td>
			Account #</td>
			<td>
			<input type="text" value="" name="account">
			</td>
			</tr>
			
			<tr>
			<td>
			Account Type
			</td>
			<td>
			<select name="acctype">
			<option value="CHECKING">Checking</option>
			</select>
					
			<select name="Orgtype1">
			<option value="Individual">Individual</option>
			<option value="Business">Business</option>
			</select>
			</td>
			</tr>
			
			<tr>
			<tr>
			<td>
			Check
			</td>
			<td>
			<input type="text" value="" name="check"><IMG 
                  src="images/echeck.jpg">
			</td>
			</tr>
			
			<tr>
			<td>
			Drivers Lic #
			</td>
			<td>
			<input type="text" value="" name="lic">
			
			<select name="drilic">
			<option>AK</option>
			<option>AL</option>
			<option>AR</option>
			<option>AZ</option>
			<option>CA</option>
			<option>CO</option>
			<option>CT</option>
			<option>DC</option>
			<option>DE</option>
			<option>FL</option>
			<option>GA</option>
			<option>HI</option>
			<option>IA</option>
			<option>ID</option>
			<option>IL</option>
			<option>IN</option>
			<option>KS</option>
			<option>KY</option>
			<option>LA</option>
			<option>MA</option>
			<option>MD</option>
			<option>ME</option>
			<option>MI</option>
			<option>MN</option>
			<option>MO</option>
			<option>MS</option>
			<option>MT</option>
			<option>NC</option>
			<option>ND</option>
			<option>NE</option>
			<option>NH</option>
			<option>NJ</option>
			<option>NM</option>
			<option>NT</option>
			<option>NV</option>
			<option>NY</option>
			<option>OH</option>
			<option>OK</option>
			<option>OR</option>
			<option>PA</option>
			<option>PR</option>
			<option>RI</option>
			<option>SC</option>
			<option>SD</option>
			<option>TN</option>
			<option>TX</option>
			<option>UT</option>
			<option>VA</option>
			<option>VT</option>
			<option>WA</option>
			<option>WI</option>
			<option>WV</option>
			<option>WY</option>
			</select>
			
			<tr>
			<td>
			Drivers Lic Exp
			</td>
			<td>			
			<select name="month" size="1">
				<option selected value="0" >Select Month</option>
				<option value="01">January</option>
				<option value="02">February</option>
				<option value="03">March</option>
				<option value="04">April</option>
				<option value="05">May</option>
				<option value="06">June</option>
				<option value="07">July</option>
				<option value="08">August</option>
				<option value="09">September</option>
				<option value="10">October</option>
				<option value="11">November</option>
				<option value="12">December</option>
			</select>
			
			<input type="text" value="" name="licexp" size="5">
            <% Current_year = year(now()) %>
			<select name="fromyy" size="1">
			   <option selected value="0">Select Year</option>
			   <% for go_year =     Current_year to Current_year + 10 %>
			         <% go_year_two_digit = right(go_year,2) %>
			         <option value="<%= go_year_two_digit  %>"><%= go_year %>
		       <% next %>
			</select>
					</td>
			</tr>
			
			<% end if%>
			<%if cardtype <> "Paypal" then%>
			<tr>
			<td>
			Amount Due
			</td>
			<td>
			$ 1.00
			</td>
			</tr>
			<tr>
			<td align=left><input type="button" border="0" value="Continue" onclick="javascript:fnshow(<%=intflag%>)">
			<input type="button" value="Back" border="0" onclick="javascript:history.back()"></td>
			<input type="hidden" name="hidflag" value="<%= intflag%>">
			<input type="hidden" name="hidtype" value="<%= cardtype%>">
			<input type="hidden" name="hidSessionID" value="<%= Session("tempresellerid")%>">
			<input type="hidden" name="AMOUNT" value="1.00">
			</tr>
			
			</form>
			<% end if%>			
<%
			if cardtype = "Paypal" then
			intflag = "3"
			
%>
	
<form name="form1" method="post" action="https://www.paypal.com/cgi-bin/webscr"> 
<input type="hidden" name="cmd" value="_ext-enter"> 
<input type="hidden" name="redirect_cmd" value="_xclick"> 
<input type="hidden" name="business" value="kurtnorway@sify.com">
<input type="hidden" name="currency_code" value="USD"> 
<input type="hidden" name="amount" value=".01"> 
<input type="hidden" name="item_number" value="<%= session("tempresellerid") %>"> 
<input type="hidden" name="item_name" value="Easystorecreator"> 
<input type="hidden" name="notify_url" value="http://admin1.easystorecreator.com/Licence_notify2.asp">
<input type="hidden" name="cancel_return" value="http://admin1.easystorecreator.com/Licence_cancel.asp">
<input type="hidden" name="return" value="http://admin1.easystorecreator.com/Licence_notify2.asp">


<table>
<tr>
<td>
<!--	<b>Click on the button below to pay now.</b><br><br>-->
</td>
</tr>
<tr>
</tr>


<input type="submit" value="Pay Now" id=submit1 name=submit1>
</form>
<% end if%>




	</table><!--#include file="footer.asp"--></BOdy>
