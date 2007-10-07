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
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="../include/commonfunctions.js" ></script>
<script language ="javascript">

//function gets executed when the user click update button
function fnUpdate(end)
{
	p1 = "0"
	p2 = document.frm.hidtemppageno.value
	
	p3 = "1"
	hitlistid = "1"    
	var ErrMsg
	ErrMsg = ""
	end = eval(end)-1
	start = eval(document.frm.hidCtrnum.value)
	
	for(i=start;i<=end;i++)
	{
		str = "document.frm.txtamt"+i+".value"
		str1 = eval(str)
		strid = "document.frm.txtamt"+i+".id"
		strid1 = eval(strid)
		strnew = "document.frm.hidbalance"+i+".value"
		str2 = eval(strnew)
		
		strnew1 = "document.frm.hidmax"+i+".value"
		str3 = eval(strnew1)
		
		strnew2 = "document.frm.hidmode"+i+".value"
		str4 = eval(strnew2)
		if (isWhitespace(str1)==false)
		{
			if( eval(str4)=='0' || eval(str4)=='')
			{
				ErrMsg = ErrMsg + "Reseller ID "+strid1+" has not chosen his payment mode.\n";
			}
			if( eval(str4)=='1' || eval(str4)=='2')
			{
			
				if (isWhitespace(str1)==false)
				{
					if (isAllNumeric(str1)==false)
					{
						ErrMsg=ErrMsg + "Please enter valid amount for Reseller ID "+strid1+".\n";
			
					}
					else if(eval(str1)==0 && eval(str2)!=0)
					{
						ErrMsg=ErrMsg + "Please enter amount greater than Zero for Reseller ID "+strid1+".\n";
					}
					
					
					else if (isAllNumeric(str1)==true)
					{

						if(eval(str1) > eval(str2))
		
							{
		
								//ErrMsg=ErrMsg + "Amount to pay should not be greater than the amount balance for Reseller ID "+strid1+" .\n";
							}
						if(eval(str1) > eval(str3))
							{
		
								//ErrMsg=ErrMsg + "Amount to pay should not be greater than then net amount for Reseller ID "+strid1+" .\n";
							}	
					}		
									
			}			
		}		
		
			
		
		}
		
		
	}

	var Flag
	Flag = "0"
	if (isAllBlank(start,end) == true)
		{
			Flag="1"		
		}
	if (Flag ==0)
	{
		ErrMsg = ErrMsg + "Please enter amount for atleasr one Reseller ID.\n"
	}	
	
	if(ErrMsg!="")
	{
			alert(ErrMsg);
	}
	
	
	if (ErrMsg=="" && Flag=="1")
		{
			document.frm.action="update_payment.asp?action=save&p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
			document.frm.submit();
		}

}
function isAllBlank(start,end)
{
	
	for(i=start;i<=end;i++)
		{
					str = "document.frm.txtamt"+i+".value"
					str1 = eval(str)
		
					if (isWhitespace(str1)== false )
					{
						return true
					}
					
		}			
}		

function fnShow(id)
{
document.frm.action = "resellers_users_list.asp?action="+id
document.frm.submit()
}

function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	document.frm.hidpagechanged.value ="True"
	document.frm.action = "update_payment.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	document.frm.submit()
	}
	
	//function for sorting 
function goSort( fldName, columnNum )
	{
	    p1 = "0"
	    p2 = "1"
	    p3 = "1"
		hitlistid = "1"    
	    theForm = document.frm;
	    
	    theForm.SortBy.value = fldName;
	    //theForm.SortMsg.value = headerText;
	    theForm.SortColumn.value = columnNum;
	    theForm.action = "update_payment.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	    theForm.submit( );
	}
	function fnOpen(reselid,month,year){


	var sUrl = "update_showbalance.asp?rid="+ reselid +"&mnth="+ month +"&yr="+ year ;
	var sTitle = "Print"
	var sOpt = "location=0,menubar=0,scrollbars=yes,top=0,resizable=1,width=700,height=300,left=25,top=20";
	window.open(sUrl, sTitle, sOpt);
		
	}
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
'code realated to sorting and all 
'******************************************************Stats here****************************************************************

if Request.Form("SortBy") <> "" then
'	2nd page navigation
	sortfield = Request.Form("SortBy")
else
	'n th page navigation
	sortfield = Request("hidsortfield")
end if	

if Request.Form("SortBy") <> "" then
	sortcol = Request("SortColumn")
else
	sortcol = Request("hidsortcol")
end if		

ordering = "" & Request("SortBy")

prefixes = Array( "","","","" )

If ordering = "" Then 
    ordering = "fld_reseller_id"
    sortColumn = 1
    'incase the user does not sort and only navigates then we need to maintain 
    'the default sort field.
    sortfield = ordering
    sortcol = sortColumn
Else
    sortColumn = CInt( Request("SortColumn") )
End If

priorOrder = "" & Request("PriorSortBy") ' what did we sort by last time?

if trim(Request.Form("hidpagechanged")) <> "" then
		
	'these values are required when no sorting and only navigation takes place.	
		ordering = Request("hidsortfield")
		sortorder = Request.Form("hidsortorder")
		sortcol = Request.Form("hidsortcol")
		ordering = ordering & " " & sortorder
		
	'	Response.Write "ordering :" & ordering & "<br>"
		if trim(sortorder)="ASC" then
			prefixes(sortcol) = "<img src=asc.gif border=0 alt='sort order ascending'>"
		else
			prefixes(sortcol) = "<img src=desc.gif border=0 alt='sort order descending'>"
		end if
		
else

	If trim(ordering) = priorOrder Then
		sortorder = "DESC"
	    ordering = ordering & " DESC" 
	    priorOrder = ""
	    prefixes(sortColumn) = "<img src=desc.gif border=0 alt='sort order descending'>"
	Else

		sortorder = "ASC"
	    priorOrder = ordering 
	    ordering = ordering & " ASC" 
	    prefixes(sortColumn) = "<img src=asc.gif border=0 alt='sort order ascending'>"
	End If
end if
'******************************************************Ends here****************************************************************


'code here to het the month and year from hidden var  or when the firs time the page is visited  or thr' quesrystring
'******************************************************Starts here****************************************************************
		'asp code here
		dim strmonth,sqlselect,rsselect,strid,stramttopaid,sqlgetbalance,rsgetdata,stramtpaid,strminamtbalance
		dim sqlpaymentmode,rsmode,stryear
		dim strMinBal, strStartDate
		'code here to reterive the values passed from the prev page
		strmonth=trim(Request.Form("frommm"))
		stryear=trim(Request.Form("fromyy"))

		'retriving value of month  from hidden variable		
		if trim(Request.Form("hidmonth")) <> "" then 
				strmonth=trim(Request.Form("hidmonth"))
		end if	
		
		'retriving value of year from hidden variable		
		if trim(Request.Form("hidyear")) <> "" then 
				stryear=trim(Request.Form("hidyear"))
		end if		
		
		if trim(Request.QueryString("checkmonth")) <> "" then 
				strmonth= trim(Request.QueryString("checkmonth"))
		end if		
		
		if trim(Request.QueryString("checkyear")) <> "" then 
				stryear= trim(Request.QueryString("checkyear"))
		end if		
		
		
		
		
		
		
'******************************************************Ends here****************************************************************
		

'code here to reterive the list of all the resellers irrespective of the month
'******************************************************Stats here****************************************************************
'Retriving the resellers list from the database
sqlGetData = "select fld_first_name+' '+fld_last_name as name,fld_reseller_id , IsNull(fld_min_amt,0) from tbl_reseller_master order by "&ordering&" "
set rsGetData=server.CreateObject("ADODB.RecordSet")
rsGetData.CursorLocation = 3


'executing sql query
rsGetData.open sqlGetData,conn,2,2
'******************************************************Ends here****************************************************************

'code here to update the payment status corresponding to all the reselles
'******************************************************Stats here****************************************************************
if trim(Request.QueryString("action")) = "save" then

	start = trim(Request.Form("hidCtrnum"))
	val =  trim(Request.Form("EndctrNum"))
	for i= start to val
		
		'retrieving the corresponding resellerid 
		id = "hidid"&i
		strresellerid=trim(Request.form(id))
		
		'retrieving the corresponding amount input 
		txtamount = "txtamt"&i
		amount = trim(Request.Form(txtamount))
		
		
		'if amount is not null then only make an insert or update 
		'*********************Starts here ***************************************
		
		if amount<> "" then
			
			'code here first to check whether there is already a record there 
			'*********************Starts here ***************************************
				strlclQC = " select fld_esc_reseller_payment_id,fld_amount_paid from TBL_ESC_RESELLER_PAYMENT WHERE fld_reseller_id ='"&strresellerid&"' and fld_month='"&trim(strmonth)&"' and fld_year='"&trim(stryear)&"'"
				set rscheckrecord = conn.execute(strlclQC)
				
			'*********************Ends here ***************************************
			
			'if the record is there then update
			if not rscheckrecord.eof then
				'code here to update 
				'*********************Starts here ***************************************
				amountpaid = trim(rscheckrecord("fld_amount_paid"))
				stramtpaid= CSng(amountpaid) + CSng(amount)
	   			'Response.Write "amountpaid "&amountpaid&"<br>" 
	   			'Response.Write "stramtpaid"&stramtpaid&"<br>" 
	   			'query for updating records
				strupdatequery= "update TBL_ESC_Reseller_Payment set fld_amount_paid='"&stramtpaid&"',fld_amount_balance=0" &_
   								" where fld_reseller_id='"&strresellerid&"' and fld_month='"&strmonth&"' and fld_Year='"&stryear&"' "   	   
 
				conn.execute strupdatequery
				
				intFlag  = "2"
				'*********************Ends here ***************************************
			
			else  'if the records is not there 
				'code to upadte 	
				'*********************Starts here ***************************************
				'if stramtpaid="" or stramtpaid="0" then
				stramtpaid= CSng(amount)
				
				sqlinsert=	"insert into tbl_esc_reseller_payment (fld_amount_paid,fld_amount_balance,fld_reseller_id,fld_month,fld_Year) values "&_
							" ('"&stramtpaid&"','0','"&strresellerid&"','"&strmonth&"','"&stryear&"')"
				conn.execute(sqlinsert)
				
				intflag = "1" 

			'end if		
				'*********************Ends here ***************************************
			end if
		end if
		'*********************Ends here ***************************************
	next
end if
'******************************************************Ends here****************************************************************

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
    <TD>
   <!--#include file="incESCmenu.asp"-->
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
          <form name="frm" method="post">
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>        
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1" id=TABLE1 width=100% >
          <tr>
  Reseller Payment Details - For <%= monthname(strmonth)%>&nbsp; <%= stryear%>
  		<hr noshade>	
          </tr>
         
		       <tr bgColor=#dddddd>
		       <td width="10%">
				<%=prefixes(1)%><a href="javascript:goSort('fld_reseller_id',1);">ID</a>
		       </td>
		                  <td width="15%">
					<strong><%=prefixes(2)%><a href="javascript:goSort('fld_first_name',2);">Resellers</strong></a>
         </td>
		
			<td width="15%">
				Net Profit
			</td>
			
			<td width="15%">
				Amount Paid
			</td>
			<td width="15%">
				Amount Balance
			</td>
			<td width="15%">
				Pay
			</td>
			<td width="15%">
				Payment Status
			</td>
			
          </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
          <%
          count=1
          if not rsGetData.eof then
	              'while not rsGetData.eof 
          ' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
												
												rsGetData.pagesize =50
												p1=Request.QueryString("p1") 
												if p1="" or isnull(p1) then
													p1=0
												end if 	
												
												if  pageno="" or isnull(pageno)  then
													pageno = 1
												elseif  pageno<=0 then
													pageno = 1 
												end if
												
												if cint(rsGetData.pagecount) < cint(pageno) then
														pageno = rsGetData.pagecount
												end if
												dim iNum
												rsGetData.absolutepage= cint(pageno)
												iNum = pageNo-1
												ctrNum = (iNum*rsGetData.PageSize)+1
												curpage=0
	                 							' ---------
					%>
					<input value="<%=ctrNum%>" type="hidden" name="hidCtrnum" >
					<%
					dim rsStartDate, rsGetTransaction, rsGetMonSum
					dim  strStartMonth, strStartYear, sqlGetTransaction
					dim intPaidAmt, intBalAmt, intGetMonthlyProfit , intGetMonthlyBalLeft, intGetSelMonSum, stramtbalance
					dim arrPaidAmt(25), arrBalAmt(25) , arrNetProfit(25)

					intGetSelMonSum = 0
					do while curpage < rsGetData.PageSize
						strfirst = trim(rsGetData("name"))
						intresellerid = trim(rsGetData("fld_reseller_id"))
						strminamtbalance	= trim(rsGetData("fld_min_amt"))				

						'Get the start date
						strStartDate = " select  top 1 fld_Plan_Transaction_Date " 
						strStartDate = strStartDate & "  from TBL_Reseller_Customer_Master where fld_reseller_id =  "& intresellerid &" "
						strStartDate = strStartDate & "  order by fld_Plan_Transaction_Date"
					'	response.write"<br>strStartDate="&strStartDate
						sqlGetSelMonSum = " select sum(fld_amount_pay_to_reseller)as sum  from tbl_reseller_customer_master  "
						sqlGetSelMonSum = sqlGetMonSum & " where month(fld_Plan_Transaction_Date) = " &strMonth  & ""
						sqlGetSelMonSum = sqlGetMonSum & " and year(fld_Plan_Transaction_Date) =" & strYear & " and fld_reseller_id = "& intresellerid &" "		

						If Not rsGetSelMonSum.Eof Then
							intGetSelMonSum = Trim(rsGetSelMonSum("sum"))
						End If

						set rsStartDate = conn.execute(strStartDate)
						If Not rsStartDate.Eof Then
							strStartDate = Trim(rsStartDate("fld_Plan_Transaction_Date"))
						End If
						strStartMonth =  Trim(month(strStartDate))
						strStartYear	= Trim(Year(strStartDate))					
					'	response.write"<br>strStartMonth="&strStartMonth
					'	response.write"<br>strStartYear="&strStartYear
'						response.write"<br>strYear="&strYear
						k = 0                                                                                    'Subscripts for dimensining the array
						For  i = strStartYear  to  strYear                                           'StartDate till the date selected
							'strStartMonth = 1						
						'If the Start Year and end year is different then  initialise strMonth=12 else take it from the prevoius form. 
							If Trim(strStartYear) <> Trim(strYear) Then
								strMonth = 12	
							Else
								strStartMonth = 1
								strMonth = trim(Request("frommm"))	
						'		response.write "<br><br>in"
							End If
					'	response.write"<br>strStartMonth="&strStartMonth
						'response.write"<br>strMonth="&strMonth
					'	response.write"<br>strStartYear="&strStartYear
						'response.write"<br>strYear="&strYear

							For j = strStartMonth to strMonth								
				'				response.write"<br>strStartMonth="&strStartMonth
						'response.write"<br>strMonth="&strMonth
				'		response.write"<br>strStartYear="&strStartYear
					

								'Calculate the amount received by customer per month starting from transaction date.								
								sqlGetMonSum = " select sum(fld_amount_pay_to_reseller)as sum  from tbl_reseller_customer_master  "
								sqlGetMonSum = sqlGetMonSum & " where month(fld_Plan_Transaction_Date) =" &j
								sqlGetMonSum = sqlGetMonSum & " and year(fld_Plan_Transaction_Date) ="& i &" and fld_reseller_id = "& intresellerid &" "												
								'response.write"<br>sqlGetMonSum="&sqlGetMonSum
								
								set rsGetMonSum = conn.execute(sqlGetMonSum)
								If Not rsGetMonSum.Eof Then
									intGetMonthlyProfit = 	Trim(rsGetMonSum("sum"))
								End If
							'	response.write"<br>intGetMonthlyProfit="&intGetMonthlyProfit & "<br>"

								'Check in the table Tbl_Esc_Payment if the amount is paid the month.
								sqlGetTransaction = "select fld_amount_paid , fld_amount_balance  "
								sqlGetTransaction = sqlGetTransaction & " from tbl_esc_reseller_payment where fld_reseller_id="& intresellerid &" "
								sqlGetTransaction = sqlGetTransaction & " and fld_month='"&strStartmonth&"' and fld_year='"&strStartyear&"' " 	
								'response.write"sqlGetTransaction="&sqlGetTransaction

								set rsGetTransaction = conn.execute(sqlGetTransaction)
								
								'intGetMonthBalNetProfit =  
								If Not rsGetTransaction.Eof Then
				'				response.write"<br><br>strStartMonth="&strStartMonth
				'				response.write"<br><br>strStartYear="&strStartYear
									intPaidAmt = Trim(rsGetTransaction("fld_amount_paid"))
									intBalAmt = Trim(rsGetTransaction("fld_amount_balance"))									
									arrPaidAmt(k)  =  intPaidAmt
									arrBalAmt(k)  =  intBalAmt
									arrNetProfit(k) =  intGetMonthlyProfit					
								'	response.write"<br><br>intPaidAmt="&intPaidAmt
								'	response.write"<br><br>intBalAmt="&intBalAmt
									'If intBalAmt = 0 Then						
									'	intGetMonthlyBalLeft = intGetMonthlyBalLeft + intBalAmt
									'Else
									'Check if the amount is zero or not.If not carry the amount to next month .
									'	If intPaidAmt >  intGetMonthlyProfit Then 											
											If cdbl(intPaidAmt) >  cdbl(intGetMonthlyProfit) Then
												intDistAmt = intDistAmt + intPaidAmt -  intGetMonthlyProfit
												arrBalAmt(k) =  0
									'			intGetMonthlyBalLeft = intGetMonthlyBalLeft + intBalAmt
												'Redimension the array
												Redim arrBalAmt(k)
												Redim arrPaidAmt(k)
												Redim arrNetProfit(k)
												'Distribute the money
												For m = 0 to Ubound(arrBalAmt) 
													intBalLeft = arrBalAmt(m)
													If Cdbl(intDistAmt) < Cdbl(arrBalAmt(m)) Then																											
														arrBalAmt(m) = arrBalAmt(m) -  intDistAmt
													'	intDistAmt = intDistAmt - intBalLeft
														intDistAmt = 0
									'					intGetMonthlyBalLeft(K) = intGetMonthlyBalLeft(K) - intDistAmt
													Else														
												'		response.write"<br>========arrBalAmt(m)="&arrBalAmt(m)
									'					intGetMonthlyBalLeft(K) = intGetMonthlyBalLeft(K) - arrBalAmt(m)
														arrBalAmt(m) = 0														
														intDistAmt = intDistAmt - intBalLeft
													End If																
												Next
											Else
												'arrBalAmt(k) = arrPaidAmt(k) - arrBalAmt(k) 
												arrBalAmt(k) = arrNetProfit(k)  - arrPaidAmt(k)   
										'		response.write"<br>***********arrbalAmt="&arrBalAmt(k)
											End If
									'	End If
									'End If					
								Else
									arrPaidAmt(k)  =  0
									arrBalAmt(k)  =  intGetMonthlyProfit
									arrNetProfit(k) =  intGetMonthlyProfit
									'intGetMonthlyBalLeft(K) = intGetMonthlyBalLeft(K) + intGetMonthlyProfit
								'	intGetMonthlyBalLeft = intGetMonthlyBalLeft + intDistAmt	
								
								End If								
								k = k + 1
								strStartMonth = strStartMonth + 1								
							Next										
							strStartYear = strStartYear +1
						Next

					'	response.write"<br>strStartDate="&strStartDate					
					intBalAmtToBePaid = 0
					intCheckBalAmtToBePaid = 0
					'response.write"<br><br>k="&k
					for i=0 to k
					'	response.write"<br><br>balCal="&arrBalAmt(i)			
						if i < k-1 Then
							intCheckBalAmtToBePaid =  intCheckBalAmtToBePaid + cdbl(arrBalAmt(i))
						End If
						intBalAmtToBePaid   =  intBalAmtToBePaid + cdbl(arrBalAmt(i))
						arrBalAmt(i) = ""
					Next
			'	response.write"<br>intBalAmtToBePaid="&intCheckBalAmtToBePaid &"<br>"
			'		response.write"<br>intBalAmtToBePaid="&intBalAmtToBePaid &"<br>"
				'	response.end
				'	response.write"<br>intGetSelMonSum="&intGetSelMonSum
				'	response.write"<br>strminamtbalance="&strminamtbalance
					If Cdbl(intGetSelMonSum) > Cdbl(strminamtbalance)  or Cdbl(intGetSelMonSum) = Cdbl(strminamtbalance)  Then
		  %>
		 
          <tr>
          	<td width="15%">
          <%=rsGetData("fld_reseller_id")%>
          </td>
          	<td width="15%">
			<%= strfirst%>
		  </td>
			
			<%
			'Retriving the net payable amount of the reseller
			
			stramount = "select sum(fld_amount_pay_to_reseller)as sum from tbl_reseller_customer_master where month(fld_Plan_Transaction_Date) ="&strmonth&" and year(fld_Plan_Transaction_Date) ="&stryear&" and fld_reseller_id="&intresellerid&""
			set rsamount =conn.execute(stramount)
			'Response.Write "stramount"&stramount
			
				amount = "0.00"
				if not rsamount.eof then
						if isNull(trim(rsamount("sum"))) = False then 
							amount = trim(rsamount("sum"))
							test = split(amount,".")
							if test(1) = "" then
								amount=amount&".00"	
							end if
						end if
				rsamount.close
				rsamount=nothing
				end if%>
				
				<td width="15%">
				$<%= formatnumber(amount,2)%> 				
			</td>
				<%
				'reteriving the amount paid to the reseller 
				sqlgetbalance="select fld_reseller_id,fld_amount_paid as paidamt from tbl_esc_reseller_payment where fld_reseller_id="&intresellerid&" and fld_month='"&strmonth&"' and fld_year='"&stryear&"' " 
				set rsgetbalance=conn.execute(sqlgetbalance) 
				'Response.Write sqlgetbalance
				
				if not rsgetbalance.eof then 	
					stramtpaid=trim(rsgetbalance("paidamt"))
					stramtbalance = CSng(amount)-CSng(stramtpaid)
					stramtbalance = formatnumber(stramtbalance,2)
					stramtpaid = formatnumber(stramtpaid,2)
				else
					stramtpaid = "0.00"
					stramtbalance = amount
				end if
					rsgetbalance.close
					set rsgetbalance=nothing
				
				
				%>
				<td width="15%">
			$<%=stramtpaid%>
			</td>
				<td width="15%">
				<%'=stramtbalance%>
			<%	If cdbl(intCheckBalAmtToBePaid ) <> 0 Then%>
					$<a href="javascript:fnOpen(<%=intresellerid%>,<%=strMonth%>,<%=strYear%>);"><%=formatnumber(intBalAmtToBePaid,2)%></a>
			<%Else%>
				<%=formatnumber(intBalAmtToBePaid,2)%>
			<%End If%>
			</td>
			<td>
				$<input type="text" name="txtamt<%=ctrNum%>" id=<%=rsGetData("fld_reseller_id")%> maxlength=6 size=6>
				<input type="hidden" value="<%=rsGetData("fld_reseller_id")%>" name="hidid<%=ctrNum%>">
			</td>
				<td width="15%">
          <%
          if amount="0.00" then
			Response.Write "No Payment"
          elseif trim(stramtpaid)= trim(stramtbalance) or trim(stramtpaid)= trim(amount)then
			Response.Write "Completed"
          else
			Response.Write "Pending"
		  end if
          %>
          </td>                   
           	<td width="15%"></td>
		</tr>
		  <%End If%>
          <tr>
				<td width="15%">

			</td>
			
			
			<td>
			<%
			'code here to retrive the payment mode os the reseller
					'*****************************************************************************************************
					sqlMode = "select fld_payment_mode from tbl_ESC_reseller_payment_mode where fld_reseller_id="&intresellerid&"" 
					set rsMode = conn.execute(sqlMode)
					if not rsMode.eof then
						mode = trim(rsMode("fld_payment_mode"))
					else
						mode = 0
					end if
					'*****************************************************************************************************
			%>		
			<input type="hidden" name="hidpaid<%=ctrNum%>" value="<%=stramtpaid%>">
			<input type="hidden" name="hidbalance<%=ctrNum%>" value="<%=stramtbalance%>">
			<input type="hidden" name="hidmax<%=ctrNum%>" value="<%=amount%>">
			<input type="hidden" name="hidmode<%=ctrNum%>" value="<%=mode%>">
           
		<%	
				
			
				ctrNum = ctrNum + 1
				curpage=curpage+1	
				
			
				
			if not rsGetData.EOF then
				rsGetData.movenext	
			end if	
			if rsGetData.EOF then
				exit do
			end if
				loop	
			
			end if	
			  %>
		</tr>	
		  </tr>
           <tr>
			<td >&nbsp;
			</td>
          </tr> 
        
        <tr>
				<td colspan="7"><input type="button" name="Update" value="Update" onclick="javascript:fnUpdate('<%=ctrNum%>')">&nbsp;
				<input type="Reset" name="Reset" value="Reset"> &nbsp;<input type="button" name="back" value="Back" onclick="javascript:history.back()"> 
				</td>	
				
		  </tr>
          	<tr>
		<td>
				<br>
				           	
				</td>	
	     <tr>         
	     		<!--<br>SortBy--><INPUT Type="hidden" Name="SortBy">
		<!--<br>SortColumn--><INPUT Type="hidden" Name="SortColumn">
		 <input type="hidden" name="EndctrNum" value="<%=ctrNum-1%>">
		<!--<br>PriorSortBy--><INPUT Type="hidden" Name="PriorSortBy" Value="<%=priorOrder%>">
		<!--<br>hidsortfield--><INPUT Type="hidden" Name="hidsortfield" value="<%=sortfield%>"> <!--second page and n th page navigation-->
		<!--<br>hidsortcol--><INPUT Type="hidden" Name="hidsortcol" value="<%=sortcol%>"> <!--second page and n th page navigation-->
		<!--<br>hidsortorder--><INPUT Type="hidden" Name="hidsortorder" Value="<%=sortorder%>"> <!--second page and n th page navigation-->
		<!--<br>hidpagechanged--><INPUT Type="hidden" Name="hidpagechanged" Value="">
		<input type="hidden" name="hidmonth" value="<%=strmonth%>">
		<input type="hidden" name="hidyear" value="<%=stryear%>">	
		<input type="hidden" name="hidtemppageno" value="<%=pageno%>">	

		
          </table></FORM>  
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
        
			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsGetData.PageCount,10 %></font></td>
		
          </tr> 
        <% rsGetData.Close
        rsGetData=nothing%>
              <TBODY>
              <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE>
</BODY></HTML>

<% 
sub DisplayNavBar(p1,CurrentPage,MaxPages,DisplayPageCount)
hitlistid=1
	p3=1
	if Cint(CurrentPage) mod cint(DisplayPageCount) = 0 then 
	
		CounterStart=CurrentPage-(cint(DisplayPageCount)-1)
	else
		CounterStart=CurrentPage-(CurrentPage mod cint(DisplayPageCount))+1

	end if
		CounterEnd=CounterStart+(cint(DisplayPageCount)-1)

	if Cint(CounterEnd) > Cint(MaxPages) then
			CounterEnd = MaxPages
	end if
	
	if Cint(CounterStart) > 1 then
			Response.Write "<font > | </font>"
			Response.Write "<a href='javascript:jsHitListAction("& p1 &",1," & p3 &"," & hitlistid &")'><font size=2 face=arial >First</font></a>"
			Response.Write " <font face=arial size=2>|</font> "
			Response.Write "<a href='javascript:jsHitListAction("& p1 &"," & CounterStart-1 & "," & p3 &"," & hitlistid &")'><font size=2 face=arial >Previous</font></a>"				
	end if				
				
	if Cint(MaxPages) > 1 then
		for counter=CounterStart to CounterEnd
			Response.Write "<font > | </font>"
	if Cint(counter) <> Cint(CurrentPage) then
			Response.Write "<font face=arial size=2 ><a href='javascript:jsHitListAction("& p1 &"," & Counter & "," & p3 &"," & hitlistid &")'>"&counter&"</font></a>"
	else
			Response.Write "<b><font face=arial size=2><font color=red >" & counter &"</font></font></b>"
	end if	
					
	if Cint(counter)<> Cint(CounterEnd) then
			Response.Write " "
	end if		
	next		
	if Cint(CounterEnd)<>Cint(MaxPages) then
		Response.Write "<font face=arial size=2> | </font>"
		Response.Write "<font face=arial size=2><a href='javascript:jsHitListAction("& p1 &"," & CounterEnd+1 & "," & p3 &"," & hitlistid &")'><font size=2 face=arial >Next</font></a></font>"
		Response.Write "<font face=arial size=2> |</font> "		
		Response.Write "<font face=arial size=2><a href='javascript:jsHitListAction("& p1 &"," & maxpages & "," & p3 &"," & hitlistid &")'><font size=2 face=arial >Last</font></a></font>"
		Response.Write "<font face=arial size=2> |</font> "				
	'	end if		
	end if
end if			
end sub

%>



<%if intflag = "1" then 
temppageno =trim(Request.Form("hidtemppageno"))
%>
	<script language="javascript">
		alert("Payment status has been added.");
		document.location.href =  "update_payment.asp?checkmonth=<%=strmonth%>&checkyear=<%=stryear%>&p2=<%=temppageno%>"
	</script>
<% end if%>
<%if intflag = "2" then 
temppageno =trim(Request.Form("hidtemppageno"))


%>

	<script language="javascript">
		alert("Payment status has been updated.");
		document.location.href = "update_payment.asp?checkmonth=<%=strmonth%>&checkyear=<%=stryear%>&p2=<%=temppageno%>"
		
		
	</script>
<% end if%>

