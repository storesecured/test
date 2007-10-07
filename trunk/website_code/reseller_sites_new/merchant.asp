<HTML>
<HEAD>
<TITLE>Best Merchant Account</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="style.css">

<!--#include file="header.asp"-->
<!--content start-->
<%
on error goto 0
if request.form("Update") <> "" then
	Set conn_store = Server.CreateObject("ADODB.connection")
	Store_Database="Ms_Sql"
	conn_store.Provider = "MSDASQL"
	
	db_Driver = "SQL Server" 'DB Server Driver Type (Now, only SQL Server)
	db_Server = "localhost" 'DB Server IP or Alias
	db_Database =	"Wizard"			'DB Name
	db_UIN		=	"melanie"	 'DB user name
	db_pwd		=	"tom237"	 'DB user password
	
	db_ConnectionString = "Driver=" & db_Driver & ";Server=" & db_Server & ";UID=" & db_UIN & ";PWD=" & db_pwd & ";Database=" & db_Database & ";"
	
	conn_store.Open db_ConnectionString
	
	Set rs_store = Server.CreateObject("ADODB.Recordset")

	sError=""
	if not isNumeric(request.form("monthly_sales")) then
		Monthly = 1000
		sError = sError &"<BR>The amount you entered for monthly sales was not numeric, using default value $1000."
	else
		Monthly = request.form("monthly_sales")
	end if

	if not isNumeric(request.form("average_sale")) then
		Average = 10
		sError = sError &"<BR>The amount you entered for average sale was not numeric, using default value $10."
	else
		Average = request.form("average_sale")
	end if

	if not isNumeric(request.form("spread_app")) then
		spread = 12
		sError = sError &"<BR>The number of months to spread the application fee was not numeric, using default value 12."
	else
		spread = request.form("spread_app")
	end if
	sql_monthly = "Get_Monthly_Gateway "&monthly&","&average&","&spread

	rs_Store.open sql_monthly,conn_store,1,1
	
	LowestTotal = 999999
	Lowest = "None"

	SecondLowestTotal = LowestTotal
	SecondLowest = Lowest

	ThirdLowestTotal = LowestTotal
	ThirdLowest = Lowest
	
	DoNotAllowRedirect= request.form("Do_Not_Allow_Redirect")
	UsCompany=request.form("us_based")
	CanadianMerchant = request.form("canada_based")
	MerchantAccount = request.form("merchant_account")
	PhoneSupport=request.form("phone_support")
	CreditScore=request.form("credit_score")
	EChecks=request.form("e_checks")


	str = str & "<table width=525><form action=""http://www.authorizenet.com/launch/ips_main.php"" name=""EMS"" method=""Post"">"
	Do while not rs_store.eof
		Non_Match=""
		Total_Monthly = rs_Store("Total_Monthly")
		if rs_Store("Extra_Tran_Fee") > 0 then
			Total_Monthly = Total_Monthly + rs_Store("Extra_Tran_Fee")
		end if
		Name = rs_Store("Gateway")
		
		if DoNotAllowRedirect <> "" and rs_Store("Shopper_Leaves") then
			Non_Match = Non_Match & "<BR>Shopper will be redirected to another site for payment processing."
		end if

		if (UsCompany <> "" or CanadianMerchant <> "") and rs_Store("US_Based_Only") then
			Non_Match = Non_Match & "<BR>Only available to US based merchants."
		end if
		
		if PhoneSupport <> "" and not rs_Store("Phone_Support") then
			Non_Match = Non_Match & "<BR>Phone support is not included with this gateway."
		end if
		
		if CreditScore = "" and rs_Store("Credit_Score") then
			Non_Match = Non_Match & "<BR>In most cases your credit score must be above 600 to qualify."
		end if
		
		if CanadianMerchant = "" and rs_Store("Canadian_Currency") then
			Non_Match = Non_Match & "<BR>Rates listed here are for Canadian Currency and do not apply to stores based outside Canada."
		end if
		
		if MerchantAccount <> "" and rs_Store("Without_Merchant_Account") then
			Total_Monthly = Total_Monthly-10
		end if
	
		if EChecks <> "" and not rs_Store("Echeck_Accepted") then
			Non_Match = Non_Match & "<BR>ECheck support is not currently provided by this gateway."
		end if
		Total_Monthly = cdbl(FormatNumber(Total_Monthly,2))
		Name_Link = "<a href=#"&Name&">"&Name&"</a>"
		if cdbl(Total_Monthly) < cdbl(ThirdLowestTotal) and Non_Match="" then
			if cdbl(Total_Monthly) < cdbl(LowestTotal) then
				ThirdLowest=SecondLowest
				ThirdLowestTotal=SecondLowestTotal
				SecondLowest = Lowest
				SecondLowestTotal=LowestTotal
				Lowest = Name_Link
				LowestTotal = Total_Monthly
			elseif cdbl(Total_Monthly) < cdbl(SecondLowestTotal) then
				ThirdLowest=SecondLowest
				ThirdLowestTotal=SecondLowestTotal
				SecondLowest = Name_Link
				SecondLowestTotal = Total_Monthly
			else
				ThirdLowest = Name_Link
				ThirdLowestTotal = Total_Monthly
			end if
		end if

		if Non_Match = "" then
			sFont = "Black"
		Else
			sFont ="Gray"
		end if
		str = str &  "<TR><TD colspan=3 width='100%'><a name="&Name&"><B><font color="&sFont&">"&Name&" Total Monthly = $"& FormatNumber(Total_Monthly,2) & "</b></a></td></tr>"
		str = str &  "<TR><TD width='5%'></TD><TD nowrap><font color="&sFont&">Application Fee = $"& rs_Store("Retail_App") & "</td><td rowspan=4 valign=top><font color="&sFont&">"&Non_Match&"<BR><BR><a href='"&rs_Store("More_Info_Link")&"'>Find out more about "&Name&" or apply for an account</a><BR><BR><BR>"&rs_Store("Special_Notes")&"</td></tr>"
		str = str &  "<TR><TD width='5%'></TD><TD nowrap><font color="&sFont&">Monthly Fee = $"& rs_Store("Retail_Monthly") & "</td></tr>"
		str = str &  "<TR><TD width='5%'></TD><TD nowrap><font color="&sFont&">Discount Rate = "& rs_Store("Retail_Discount")*100 &"%</td></tr>"
		str = str &  "<TR><TD width='5%'></TD><TD nowrap><font color="&sFont&">Transaction Fee = $"& FormatNumber(rs_Store("Retail_Trans"),2) & "</td></tr>"
		if rs_Store("Trans_Limit") <> 0 then
			str = str &  "<TR><TD width='5%'></TD><TD nowrap><font color="&sFont&">Transactions over "&rs_Store("Trans_Limit") &" = $ "& FormatNumber(rs_Store("Extra_Fee"),2) & " extra</td></tr>"
		end if
		str = str &  "<TR><TD colspan=4><HR></TD></tr>"

		rs_store.movenext
	loop
	rs_store.close

	if sError <> "" then
		response.write "<table width=525><tr><td><font color=red>"&sError&"</font></td></tr></table>"
	end if
	if Lowest <> "None" then
		response.write "<table width=525><tr><td>The 3 lowest cost providers that match your criteria are "&Lowest&", "&SecondLowest&" and "&ThirdLowest&".  With prices of $"&LowestTotal & ",$"& SecondLowestTotal & " and $" & ThirdLowestTotal & " respectively.  See below for pricing details and signup links.</td></tr></table>"
	else
		response.write "<table width=525><tr><td>No providers were found which matched your criteria exactly.  Please try loosening you criteria or see below for pricing details and signup links.</td></tr></table>"
	end if
	str = str & "</form></table>"
	response.write str
else
%>
<form method="POST" action="best-merchant-account.asp" name="BestMerchant">
<table width=525 align=left cellpadding=5>

		<tr><td>What is your average monthly sales?</td><td><input type=text name="monthly_sales" size=5></td></tr>
		<tr><td>What is your average sale?</td><td><input type=text name="average_sale" size=5></td></tr>
		<tr><td>How many months should we spread the application fee over?</td><td><input type=text name="spread_app" size=5> months</td></tr>
		<tr><td>Shoppers must stay at my site during checkout?</td><td><input type="checkbox" name="Do_Not_Allow_Redirect"></td></tr>
		<tr><td>Do you already have a merchant account?</td><td><input type="checkbox" name="merchant_account"></td></tr>
		<tr><td>I am not based in the US?</td><td><input type="checkbox" name="us_based"></td></tr>
		<tr><td>Are you based in Canada?</td><td><input type="checkbox" name="canada_based"></td></tr>
		<tr><td>Do you require phone support?</td><td><input type="checkbox" name="phone_support"></td></tr>
		<tr><td>Is your personal credit score > 600?</td><td><input type="checkbox" name="credit_score"></td></tr>
		<tr><td>Do you require E-Check processing?</td><td><input type="checkbox" name="e_checks"></td></tr>
		<tr><td colspan=2><input type="submit" name=Update value="Compute"></td></tr>

</table>
</form>

<% end if %>
     <!--content end-->
<!--#include file="footer.html"-->
