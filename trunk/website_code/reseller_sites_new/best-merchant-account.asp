<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="best-merchant-account"

%>
<!--#include file="header.asp"-->

<%
on error goto 0
if request.form("Update") <> "" then
	Set conn_store = Server.CreateObject("ADODB.connection")
	Store_Database="Ms_Sql"
	'conn_store.Provider = "MSDASQL"
	conn_store.Provider = "sqloledb.1"

	db_Driver = "SQL Server" 'DB Server Driver Type (Now, only SQL Server)
	'db_Server = "69.20.16.229,1433" 'DB Server IP or Alias
	db_Server = "localhost,1433" 'DB Server IP or Alias
	'db_Database =	"Wizard" 		'DB Name
	db_UIN		=	"melanie"	 'DB user name
	db_pwd		=	"tom237"  'DB user password

	db_ConnectionString = "Provider=sqloledb;Data Source=" & db_Server & ";UID=" & db_UIN & ";PWD=" & db_pwd & ";Initial Catalog=" & db_Database & ";Network Library=DBMSSOCN"

	conn_store.Open db_ConnectionString

	Set rs_store = Server.CreateObject("ADODB.Recordset")

	sError=""
	if not isNumeric(request.form("monthly_sales")) then
		Monthly = 1000
		sError = sError &"<BR>The amount you entered for monthly sales was not numeric, you must enter a valid number, using default value $1000."
	elseif request.form("monthly_sales") <= 0 then
                Monthly = 1000
		sError = sError &"<BR>The amount you entered for monthly sales was invalid, you must enter a number greater than 0, using default value $1000."
        else
		Monthly = request.form("monthly_sales")
	end if

	if not isNumeric(request.form("average_sale")) then
		Average = 10
		sError = sError &"<BR>The amount you entered for average sale was not numeric, you must enter a valid number, using default value $10."
	elseif request.form("average_sale") <= 0 then
		Average = 10
		sError = sError &"<BR>The amount you entered for average sale was invalid, you must enter a number greater than 0, using default value $10."
        else
		Average = request.form("average_sale")
	end if

	if not isNumeric(request.form("spread_app")) then
		spread = 12
		sError = sError &"<BR>The number of months to spread the application fee was not numeric, you must enter a valid number, using default value 12."
	elseif request.form("spread_app") <= 0 then
		spread = 12
		sError = sError &"<BR>The number of months to spread the application fee was invalid, you must enter a number greater than 0, using default value 12."

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
	Location = request.form("location")
	HighRisk = request.form("high_risk")


	MerchantAccount = request.form("merchant_account")
	PhoneSupport=request.form("phone_support")
	CreditScore=request.form("credit_score")
	EChecks=request.form("e_checks")


	str = str & "<table width=525><form action=""http://www.authorizenet.com/launch/ips_main.php"" name=""EMS"" method=""Post""><input type=""hidden"" name=""anUserID"" value=""easystorecreateapp"">"
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

		if Location = 1 and not rs_Store("Accept_US")  then
			Non_Match = Non_Match & "<BR>The gateway does not support US merchants."
		end if
		
		if Location = 2 and not rs_Store("Accept_Canadian")  then
			Non_Match = Non_Match & "<BR>The gateway does not support Canadian merchants."
		end if
		
		if PhoneSupport <> "" and not rs_Store("Phone_Support") then
			Non_Match = Non_Match & "<BR>Phone support is not included with this gateway."
		end if
		
		if CreditScore = "" and rs_Store("Credit_Score") then
			Non_Match = Non_Match & "<BR>In most cases your credit score must be above 600 to qualify."
		end if
		
		if Location <> 2 and rs_Store("Canadian_Currency") then
			Non_Match = Non_Match & "<BR>Rates listed here are for Canadian Currency and do not apply to stores based outside Canada."
		end if
		
		if Location=3 and not rs_store("Accept_International") then
			Non_Match = Non_Match & "<BR>This gateway does not accept merchants outside the US and Canada."
		end if
		
		if HighRisk<>"" and not rs_store("Accept_HighRisk") then
			Non_Match = Non_Match & "<BR>This gateway does not accept high risk or adult oriented merchants."
		end if
		
		'if MerchantAccount <> "" and rs_Store("Without_Merchant_Account") then
		'	Total_Monthly = Total_Monthly-10
		'end if
	
		if EChecks <> "" and not rs_Store("Echeck_Accepted") then
			Non_Match = Non_Match & "<BR>ECheck support is not currently provided by this gateway."
		end if
		
		if Non_Match <> "" then
			Non_Match = "This gateway has been greyed out because they do not meet the following requested criteria:<BR>"&Non_Match
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
		response.write "<table width=525><tr><td><BR><font size=2>The 3 lowest cost providers that match your criteria are "&Lowest&", "&SecondLowest&" and "&ThirdLowest&".  With prices of $"&LowestTotal & ",$"& SecondLowestTotal & " and $" & ThirdLowestTotal & " respectively.  See below for pricing details and signup links.</font></td></tr>"
		response.write "<tr><td colspan=3><BR><B>If you are new to ecommerce please click here for more information about <a href='definitions.asp'>payment gateways and merchant accounts</a>.</b></td></tr>"
		response.write "<tr><td colspan=3><BR>If you have any questions about the providers choosen for you or don't quite understand the results please contact us before choosing a provider.  We have merchant account representatives who are trained to help you make an informed decision.</td></tr></table>"
	else
		response.write "<table width=525><tr><td><BR>No providers were found which matched your criteria exactly.  Please try loosening you criteria or see below for pricing details and signup links.</td></tr>"
		response.write "<tr><td colspan=3><BR>If you are having trouble finding a provider who meets your companies needs please contact us.  We have merchant account representatives who are trained to help you make an informed decision and find the best merchant account for your business.</td></tr></table>"
	end if
	str = str & "</form></table>"
	response.write str
else
%>
<form method="POST" action="best-merchant-account.asp" name="BestMerchant">
<table width=525 align=left cellpadding=5>
	<tr><td colspan=3><B>If you are new to ecommerce please click here for more information about <a href="definitions.asp">payment gateways and merchant accounts</a> or see a <a href=merchant-fees.asp>table of rates</a>.</b></td></tr>

      <tr><td colspan=3><B>Lowest Rate Calculator</b></td></tr>
		<tr><td>What is your average monthly sales?</td><td><input type=text name="monthly_sales" size=5></td></tr>
		<tr><td>What is your average sale?</td><td><input type=text name="average_sale" size=5></td></tr>
		<tr><td>Spread the application fee over?</td><td><input type=text name="spread_app" size=5> months</td></tr>
		<tr><td>I don't want shoppers to ever leave my site?</td><td><input type="checkbox" name="Do_Not_Allow_Redirect"></td></tr>
		<tr><td>I already have a merchant account?</td><td><input type="checkbox" name="merchant_account"></td></tr>
		<tr><td>I am based in the US</td><td><input type="radio" name="location" value=1 checked></td></tr>
		<tr><td>I am based in Canada</td><td><input type="radio" name="location" value=2></td></tr>
		<tr><td>I am based outside the US and Canada</td><td><input type="radio" name="location" value=3></td></tr>
		<tr><td>I require phone support?</td><td><input type="checkbox" name="phone_support"></td></tr>
		<tr><td>My personal credit score is > 600?</td><td><input type="checkbox" name="credit_score"></td></tr>
		<tr><td>I require E-Check processing?</td><td><input type="checkbox" name="e_checks"></td></tr>
		<tr><td>My business is considered adult oriented or high risk?</td><td><input type="checkbox" name="high_risk"></td></tr>
		<tr><td colspan=2><input type="submit" name=Update value="Compute"></td></tr>

</table>
</form>

<% end if %>

<!--#include file="footer.asp"-->
