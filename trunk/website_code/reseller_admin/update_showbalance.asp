<!--#include file = "include/ESCheader.asp"-->


<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the amount payable ,amount due to the resellor.
'	Page Name:		   	update_payment.asp
'	Version Information:EasystoreCreator
'	Input Page:		    resellers_reports.asp
'	Output Page:	    update_payment.asp
'	Date & Time:		13 Aug - 14 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
msel7=3
msel8=2
%>

<html>
<head>
<title> New Document </title>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="../include/commonfunctions.js" ></script>

</head>
<%
dim rsStartDate, rsGetTransaction, rsGetMonSum
dim  strStartMonth, strStartYear, sqlGetTransaction
dim intPaidAmt, intBalAmt, intGetMonthlyProfit , intGetMonthlyBalLeft, intGetSelMonSum, stramtbalance
dim arrPaidAmt(25), arrBalAmt(25) , arrNetProfit(25)
 arr = array(3,1)
intresellerid =Trim(Request("rid"))
strMonth = Trim(Request("mnth"))
'response.write"<br>strMonth="&strMonth
strYear = Trim(Request("yr"))
intGetSelMonSum = 0
sqlGetData = "select fld_first_name+' '+fld_last_name as name,fld_reseller_id , IsNull(fld_min_amt,0) as fld_min_amt from tbl_reseller_master where fld_reseller_id="&intresellerid

set rsGetData=server.CreateObject("ADODB.RecordSet")							
rsGetData.CursorLocation = 3
rsGetData.open sqlGetData,conn,2,2


strfirst = trim(rsGetData("name"))
strminamtbalance	= trim(rsGetData("fld_min_amt"))				
'Get the start date
strStartDate = " select  top 1 fld_Plan_Transaction_Date, datediff(month, fld_Plan_Transaction_Date,'"&strYear&"-"& strMonth &"-01 00:00:00.000') months " 
strStartDate = strStartDate & "  from TBL_Reseller_Customer_Master where fld_reseller_id =  "& intresellerid &" "
strStartDate = strStartDate & "  order by fld_Plan_Transaction_Date"

'response.write "strStartDate" & strStartDate 

sqlGetSelMonSum = " select sum(fld_amount_pay_to_reseller)as sum  from tbl_reseller_customer_master  "
sqlGetSelMonSum = sqlGetSelMonSum & " where month(fld_Plan_Transaction_Date) = " &strMonth  & ""
sqlGetSelMonSum = sqlGetSelMonSum & " and year(fld_Plan_Transaction_Date) =" & strYear & " and fld_reseller_id = "& intresellerid &" "		

'response.write "<br>sqlGetSelMonSum=" & sqlGetSelMonSum
set rsGetSelMonSum = conn.execute(sqlGetSelMonSum)  

If Not rsGetSelMonSum.Eof Then intGetSelMonSum = Trim(rsGetSelMonSum("sum"))
set rsStartDate = conn.execute(strStartDate)

If Not rsStartDate.Eof Then 
	strStartDate = Trim(rsStartDate("fld_Plan_Transaction_Date"))
	redim arr(3,rsStartDate("months"))
End If

'response.write "arr " & ubound


strStartMonth =  Trim(month(strStartDate))
strStartYear	= Trim(Year(strStartDate))			
Dim k
k = 0                                                                                    'Subscripts for dimensining the array

For  i = strStartYear  to  strYear                                           'StartDate till the date selected
	If Trim(strStartYear) <> Trim(strYear) Then
		strMonth = 12	
	Else
		strStartMonth = 1
		strMonth = Trim(Request("mnth"))
	End If
	For j = strStartMonth to strMonth		
				'response.write"<br>strStartMonth="&strStartMonth
				'response.write"<br>strMonth="&strMonth
				'response.write"<br>strStartYear="&strStartYear
		'Calculate the amount received by customer per month starting from transaction date.								
		sqlGetMonSum = " select sum(fld_amount_pay_to_reseller)as sum  from tbl_reseller_customer_master  "
		sqlGetMonSum = sqlGetMonSum & " where month(fld_Plan_Transaction_Date) =" &j
		sqlGetMonSum = sqlGetMonSum & " and year(fld_Plan_Transaction_Date) ="& i &" and fld_reseller_id = "& intresellerid &" "												

		set rsGetMonSum = conn.execute(sqlGetMonSum)
		If Not rsGetMonSum.Eof Then
			intGetMonthlyProfit = 	Trim(rsGetMonSum("sum"))
		End If

		'Check in the table Tbl_Esc_Payment if the amount is paid the month.
		sqlGetTransaction = "select fld_amount_paid , fld_amount_balance  "
		sqlGetTransaction = sqlGetTransaction & " from tbl_esc_reseller_payment where fld_reseller_id="& intresellerid &" "
		sqlGetTransaction = sqlGetTransaction & " and fld_month='"&strStartmonth&"' and fld_year='"&strStartyear&"' " 	
		set rsGetTransaction = conn.execute(sqlGetTransaction)

		If Not rsGetTransaction.Eof Then
			intPaidAmt = Trim(rsGetTransaction("fld_amount_paid"))
			intBalAmt = Trim(rsGetTransaction("fld_amount_balance"))									
			arrPaidAmt(k)  =  intPaidAmt
			arrBalAmt(k)  =  intBalAmt
			arr(2,k)	 = arrBalAmt(k)'=============
			arrNetProfit(k) =  intGetMonthlyProfit					
			'Check if the amount is zero or not.If not carry the amount to next month .

			If cdbl(intPaidAmt) >  cdbl(intGetMonthlyProfit) Then
				intDistAmt = intDistAmt + intPaidAmt -  intGetMonthlyProfit
				arrBalAmt(k) =  0
				arr(2,k)	 = arrBalAmt(k) '=============
				'Redimension the array
				Redim arrBalAmt(k)
				Redim arrPaidAmt(k)
				Redim arrNetProfit(k)
				'Distribute the money
				For m = 0 to Ubound(arrBalAmt) 
					intBalLeft = arrBalAmt(m)
					If Cdbl(intDistAmt) < Cdbl(arrBalAmt(m)) Then																											
						arrBalAmt(m) = arrBalAmt(m) -  intDistAmt
					'	arr(2,k)	 = arrBalAmt(k)'==============
						intDistAmt = 0
					Else														
						arrBalAmt(m) = 0						
					'	arr(2,k)	 = arrBalAmt(k)'==========
						intDistAmt = intDistAmt - intBalLeft
					End If																
				Next
			Else	
				arrBalAmt(k) = arrNetProfit(k)  - arrPaidAmt(k)   	
				'arr(2,k)	 = arrBalAmt(k)'========
			End If
		Else
			arrPaidAmt(k)  =  0
			arrBalAmt(k)  =  intGetMonthlyProfit
			'arr(2,k)	 = arrBalAmt(k)'=====
			 arrNetProfit(k) =  intGetMonthlyProfit
		End If								
		arr(0,k)	 = strStartYear  '2004
		arr(1,k)	 = monthname(strStartMonth)  '5 
		'arr(2,k)	 = arrBalAmt(k)
		'response.write"---------------------k+1="&k
		k = k + 1
		strStartMonth = strStartMonth + 1								
	Next										
	strStartYear = strStartYear +1
Next
intBalAmtToBePaid = 0
'response.write"---------------------k="&k
for i = 0 to k
	if cstr(arrBalAmt(i)) = " " Then
		arr(2,i) = 0
	Else
		arr(2,i) = arrBalAmt(i) 
	End If
'	response.write"<br><br>balCal="&arrBalAmt(i)
'	response.write"<br><br>arr="&arr(2,i)
	intBalAmtToBePaid   =  intBalAmtToBePaid + cdbl(arrBalAmt(i))
's	arrBalAmt(i) = ""
'	response.write  "arr(0,"& i &") = " &  arr(0,i) &   "&nbsp;&nbsp;&nbsp;"
'	response.write  "arr(1,"& i &") = " &  left(arr(1,i),3) &   "&nbsp;&nbsp;&nbsp;"
'	response.write  "arr(2,"& i &") = " &  arr(2,i) &   "<br>"

Next%>

<body>
	<table cellspacing="0" cellpadding="1" border="1" align="center" width="500">
		<tr>
			<td  colspan="12" align="center"><b>Balance Details of reseller <%=strfirst%></b><td>
		</tr>
		<tr>
			<td colspan="12" >&nbsp;</td>
		</tr>
	<%For i = 0  To  k-1 %>	
		<tr>
			<td colspan="12" align="center"><b>For the Year <%= arr(0,i)%></b></td>
		</tr>
		<tr>
			<% For j = i  To  k %>
				<td align="center"><%=Left(arr(1,j),3)%></td>
				<%If arr(0,j) <> arr(0, j + 1) Then j = k 
			Next%>
		</tr>
		<tr>
			<% For j = i  To  k %>
				<td align="center"><%=formatnumber(arr(2,j),2)%></td>
				<%If arr(0,j) <> arr(0, j + 1) Then  	j = k 
				i = i + 1
			Next%>
		</tr>
		<tr>
			<td colspan="12" >&nbsp;</td>
		</tr>
	<% i = i - 1
	Next%>
	</table>

</body>
</html>
