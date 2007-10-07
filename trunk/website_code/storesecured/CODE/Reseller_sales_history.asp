<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page displays sales history of the reseller.
'	Page Name:		   	reseller_sales_history.htm
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_sales_history.asp
'	Output Page:	    reseller_sales_history.asp
'	Date & Time:		10th and 11 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel7=3
mSel8=2

startdate = trim(Request.Form("startdatebox"))
enddate = trim(Request.Form("enddatebox"))
if startdate <> "" then
sDate = split(startdate,"/")
sDay = sDate(0)
sMonth = sDate(1)
sYear = sDate(2)
qStartDate = sMonth&"/"&sDay&"/"&sYear

else
'code added to concatenate 0 with the day and month
	sDate=date()
	sDate = split(sDate,"/")
	sDay = sDate(1)
	if  sDay <10 then 
		sDay = 0&sDay
	end if 

	sMonth = sDate(0)
	if  sMonth <10 then 
		sMonth = 0&sMonth
	end if 

sYear = sDate(2)
StartDate = sMonth&"/"&sDay&"/"&sYear


'	startdate=month(date)&"/"&day(date)&"/"&year(date)


end if

if enddate <> "" then
	eDate = split(enddate,"/")
	eDay = eDate(0)
	eMonth = eDate(1)
	eYear = eDate(2)
	qEndDate = eMonth&"/"&eDay&"/"&eYear
else
'code added to concatenate 0 with the day and month
	eDate=date()
	eDate = split(eDate,"/")
	eDay = eDate(1)
	if  eDay <10 then 
		eDay = 0&eDay
	end if 

	eMonth = eDate(0)
	if  eMonth <10 then 
		eMonth = 0&eMonth
	end if 
eYear = sDate(2)
enddate = eMonth&"/"&eDay&"/"&eYear
'enddate=month(date)&"/"&day(date)&"/"&year(date)
end if

%>
<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<script src="include/script.js" language="JavaScript" type="text/javascript"></script>
<SCRIPT LANGUAGE="JavaScript" SRC="include/CalendarPopup.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
<script language="javascript">

var dtCh= "/";
	var minYear=1900;
	var maxYear=2004;
	var errflag=""; 

	function isInteger(s){
		var i;
		for (i = 0; i < s.length; i++){   
			// Check that current character is number.
			var c = s.charAt(i);
			if (((c < "0") || (c > "9"))) return false;
		}
		// All characters are numbers.
		return true;
	}

	function stripCharsInBag(s, bag){
		var i;
		var returnString = "";
		// Search through string's characters one by one.
		// If character is not in bag, append to returnString.
		for (i = 0; i < s.length; i++){   
			var c = s.charAt(i);
			if (bag.indexOf(c) == -1) returnString += c;
		}
		return returnString;
	}

	function daysInFebruary (year){
		// February has 29 days in any year evenly divisible by four,
		// EXCEPT for centurial years which are not also divisible by 400.
		return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
	}
	function DaysArray(n) {
		for (var i = 1; i <= n; i++) {
			this[i] = 31
			if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
			if (i==2) {this[i] = 29}
	} 
	return this
	}

function isDate(dtStr)
{
		var daysInMonth = DaysArray(12)
		var pos1=dtStr.indexOf(dtCh)
		var pos2=dtStr.indexOf(dtCh,pos1+1)
		var strMonth=dtStr.substring(0,pos1)
		var strDay=dtStr.substring(pos1+1,pos2)
		var strYear=dtStr.substring(pos2+1)
		strYr=strYear
		if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
		if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
		for (var i = 1; i <= 3; i++) {
			if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
		}
		month=parseInt(strMonth)
		day=parseInt(strDay)
		year=parseInt(strYr)
		if (pos1==-1 || pos2==-1){
			errflag = "1";
			//alert("The date format should be : mm/dd/yyyy")
			return false
		}
		if (strMonth.length<1 || month<1 || month>12){
			errflag = "2";		
			//alert("Please enter a valid month")
			return false
		}
		if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
			errflag = "3";
			//alert("Please enter a valid day")
			return false
		}
		if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
			errflag = "4";
			//alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
			return false
		}
		if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
			errflag = "5";
			//alert("Please enter a valid date")
			return false
		}
	return true
}
function fndisplay()
{
	var ErrMsg
	ErrMsg="";
// new code for date validation
	
if(document.frm.Start_Date.value == "")
	{
	ErrMsg = ErrMsg + "From Date is compulsory.\n"
	sDayFlag = false
	}
else if(document.frm.End_Date.value == "")
	{
	ErrMsg = ErrMsg + "To Date is compulsory.\n"
	eDayFlag = false
	}
	
else if (isDate(document.frm.Start_Date.value)==false)
			{
				switch(errflag)
				{
					case "1" :
					{
						ErrMsg = ErrMsg + "The Date format for From Date should be : mm/dd/yyyy. \n";
						break;
					}
					case "2" :
					{
						ErrMsg = ErrMsg + "Please enter a valid month for From Date. \n";
						break;
					}
					case "3" :
					{
						ErrMsg = ErrMsg + "Please enter a valid day for From Date. \n";
						break;
					}
					case "4" :
					{
						ErrMsg = ErrMsg + "Please enter a valid 4 digit year between "+minYear+" and "+maxYear+" for From Date. \n";
						break;
					}
					case "5" :
					{
						ErrMsg = ErrMsg + "Please enter a valid date for From Date. \n";
						break;
					}
					
				}	
			}
else if (isDate(document.frm.End_Date.value)==false)
			{
				switch(errflag)
				{
					case "1" :
					{
						ErrMsg = ErrMsg + "The Date format for To Date should be : mm/dd/yyyy. \n";
						break;
					}
					case "2" :
					{
						ErrMsg = ErrMsg + "Please enter a valid month for date of To Date. \n";
						break;
					}
					case "3" :
					{
						ErrMsg = ErrMsg + "Please enter a valid day for To Date. \n";
						break;
					}
					case "4" :
					{
						ErrMsg = ErrMsg + "Please enter a valid 4 digit year between "+minYear+" and "+maxYear+" for To Date. \n";
						break;
					}
					case "5" :
					{
						ErrMsg = ErrMsg + "Please enter a valid date for To Date. \n";
						break;
					}
				}	
			}
else
{
			
	var Start_Date,End_Date,tempBound,lDate,lMonth,lYear,lEndDate,lEndMonth,lEndYear,tempeBound
	var lDateToday,lMonthToday,lYearToday
	var lmsg
		Start_Date = document.frm.Start_Date.value
		End_Date = document.frm.End_Date.value
		
		tempBound = Start_Date.indexOf("/")	
		lMonth = eval(Start_Date.substring(0,tempBound))
		
		Start_Date = Start_Date.substring(tempBound+1, Start_Date.length)	
		tempBound = Start_Date.indexOf("/")	
		lDate = eval(Start_Date.substring(0,tempBound))
		
		lYear = eval(Start_Date.substring(tempBound+1, Start_Date.length))
			
		tempBound = End_Date.indexOf("/")	
		lEndMonth = eval(End_Date.substring(0,tempBound))
		
		End_Date = End_Date.substring(tempBound+1, End_Date.length)	
		tempBound = End_Date.indexOf("/")	
		lEndDay = eval(End_Date.substring(0,tempBound))
		
		lEndYear = eval(End_Date.substring(tempBound+1, End_Date.length))
		
		
		lDateToday=document.frm.hidCurrDay.value
		lMonthToday=document.frm.hidCurrMonth.value
		lYearToday=document.frm.hidCurrYear.value
		
		if (lYear != "" && lYearToday != "" && lMonth != "" && lMonthToday != "" && lDate != "" && lDateToday != "")
			{ 
				if (lYear != "" && lYearToday != "")
				{
						
					if (lYear > lYearToday)
					{
						
						flag = true
						ErrMsg = ErrMsg + ("From Date cannot be greater than today.\n")
					}
				}
	
				if (lYear != "" && lYearToday != "" && lMonth != "" && lMonthToday != "")
				{
					if (lYear == lYearToday)
					{
					
						lmonth1 = lMonth
						lmonth2 = lMonthToday
					
						if (eval(lmonth1) == eval(lmonth2))
						{
							if (lDate > lDateToday)
							{
								
								flag = true
								
								ErrMsg = ErrMsg + ("From Date cannot be greater than today.\n")												
							}
						}
						if (eval(lmonth1) > eval(lmonth2))
						{
							
							flag = true
							ErrMsg = ErrMsg + ("From Date cannot be greater than today.\n")
						}
					}	
				}		
			}
				
				
			if (lYear != "" && lEndYear != "" && lMonth != "" && lEndMonth != "" && lDate != "" && lEndDate != "")
			{ 
				
				if (lYear != "" && lEndYear != "")
				{
						
					if (lYear > lEndYear)
					{ 
						flag = true
						ErrMsg = ErrMsg + ("From Date cannot be greater than To Date.\n")
					}
				}
	
				if (lYear != "" && lEndYear != "" && lMonth != "" && lEndMonth != "")
				{
					if (lYear == lEndYear)
					{
					
						lmonth1 = lMonth
						lmonth2 = lEndMonth
					
						if (eval(lmonth1) == eval(lmonth2))
						{
							if (lDate > lEndDate)
							{
								flag = true
								ErrMsg = ErrMsg + ("From Date cannot be greater than To Date.\n")												
							}
						}
						if (eval(lmonth1) > eval(lmonth2))
						{
							flag = true
							ErrMsg = ErrMsg + ("From Date cannot be greater than To Date.\n")
						}
					}	
				}		
			}
		
}	
	if(ErrMsg !="")
		{
		alert(ErrMsg);
		}
		
		
else
		{
		
		
		document.frm.action="resellers_sales_report.asp";
		document.frm.submit();
		}
	

}

</script>

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
End_Date = FormatDateTime(now(),2)
Start_Date = dateadd("m",-1,End_Date)
dim rsselect,strgetdata,intresellerid,noofcust,profitamt,strescdata,rsescdata,strstatus,rsstatus,strstatnam,paidamt,balanceamt,currentmonth
dim sqldate,rsdate
intresellerid = session("Resellerid")
currentmonth=month(now)
currentyear=year(now)

'query for retriving no. of customers,sum of amount payable to him, month of the date according to reseller id and current month
		strgetdata="select count(distinct(fld_customer_id)) as count,sum(fld_amount_pay_to_reseller) as amount from tbl_reseller_customer_master where"&_ 
		" fld_reseller_id="&intresellerid&" and month(fld_plan_transaction_date)="&currentmonth&" and year(fld_plan_transaction_date)="&currentyear&" "
				
		set rsselect=conn.execute(strgetdata)
if not rsselect.eof then 
				noofcust=trim(rsselect("count"))
				profitamt=trim(rsselect("amount"))
				test = split(profitamt,".")
				if test(1) = "" then
					profitamt=profitamt&".00"
				end if
end if

		rsselect.close
		set rsselect=nothing

	'query for retriving amount paid to him,balance amount	
	
	strescdata="select sum(fld_amount_paid) as amountpaid,sum(fld_amount_balance) as amountbalance from tbl_esc_reseller_payment where fld_reseller_id="&intresellerid&" and "&_
			   " fld_month="&currentmonth&" and fld_year="&currentyear&" "
			   	
	set rsescdata=conn.execute(strescdata)
		
if not rsescdata.eof then
		paidamt=trim(rsescdata("amountpaid"))
		test = split(paidamt,".00")
		if test(1) = "" then
			paidamt=paidamt&".00"
		end if
		balanceamt=trim(rsescdata("amountbalance"))
		test1 = split(balanceamt,".")
		if test1(1) = "" then
			balanceamt=balanceamt&".00"
		end if
		if profitamt-balanceamt=0 then
		
			strstatname ="Complete"
		else
			strstatname	 ="Pending"
		end if
else
		paidamt="0"
		balanceamt=profitamt
		strstatname="Pending"

end if		

if noofcust=0 then
		profitamt="0"
		paidamt="0"
		balanceamt="0"
		strstatname=""

end if
		rsescdata.close
		set rsescdata=nothing

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
  <TR>
    <TD>
      <!--#include file="incmenu.asp"-->
            
            </TD></TR>
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
          <!-- CURRENT DATES -->
			<input type=hidden name="hidCurrDay" value="<%=day(date)%>">
			<input type=hidden name="hidCurrMonth" value="<%=month(date)%>">				
			<input type=hidden name="hidCurrYear" value="<%=year(date)%>">
			<input type="hidden" name="hidSubmit">
			<!-- CURRENT DATES -->

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
          <TD class=pagetitle height=400 vAlign=top>
          Current Month : <%=monthname(currentmonth)%> &nbsp;<%=currentyear%>
          <table border="1" width=100%>
          <br>
          <tr bgColor=#dddddd>
			<td  align="middle">
				Total Customers Referred
			</td>
			<td  align="middle">
				Net Profit
			</td>
			<td align="middle">
				Payment Status 
			</td>
			
			
          </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
          <tr>
          <td width="30%" align="middle">
				
			<%=noofcust%>	
			</td>
			
			<td  align="middle">
				
				$ <%=profitamt%>
			</td>
			          <td  align="middle">
				
					<%=strstatname%>
			</td>
		
          </tr>
           <td width="30%" align="middle">
		
		</td>
		  </tr>
        </table>
       <br>
       <br>   
       Amount Status<br>
      <table  border="1" width="100%" >
      <br>
      <tr bgColor=#dddddd>
		<td>
		Net Payable Amount 
		</td>
		<td>
			Amount Paid
		</td>
		<td>
			Amount Due
		</td>
      
      </tr>
        <tr>
		<td>
			$ <%=profitamt%>
		</td>
		<td>
			
			$ <%=paidamt%>
		</td>
		<td>
			
			$ <%=csng(profitamt)-csng(paidamt)%>
		</td>
      
      </tr>
      
      </table>
          <table border="1">
          <tr>
          <br>
          My Sales History
		<hr noshade>	
          </tr>
          
          <tr>
			<td>
			
			</td>
			         
			<td>
		
		<SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
		<tr bgcolor="#ffffff"> 
		<td width="18%" height="28"><font face="Arial, Helvetica, sans-serif" size="2">From Date&nbsp;</font></td>
		</td>
		<td width="83%" height="28"><input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
		 <A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="cal1.select(document.forms[0].Start_Date,'anchor2','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
		 </td>
		<tr bgcolor="#ffffff"> 
				<td width="18%" height="28"><font face="Arial, Helvetica, sans-serif" size="2">To Date&nbsp;</font></td>
		</td>
		<td width="83%" height="28"><input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
		<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>			
		</td>
	</tr>
	</tr>						
		<!--CODE COMMENTED HERE 	<tr bgcolor="#ffffff"> 
			<td width="18%" height="28"><font face="Arial, Helvetica, sans-serif" size="2">From Date&nbsp;</font></td>
			<td width="83%" height="28"><INPUT onclick="javascript:alert('Click on calendar image.')" 
							readOnly size=15  value="<%=startdate%>"name=startdatebox>
							<A onmouseover="window.status='Date Picker';return true;" 
							onmouseout="window.status='';return true;" 
							href="javascript:showCal('StartDate');">
							<IMG height=20 	alt="Select from date." src="images/calendar.gif" width=30 border=0></td>
		</tr>
		
		
		<tr bgcolor="#ffffff"> 
			<td width="18%" height="28"><font face="Arial, Helvetica, sans-serif" size="2">To Date&nbsp;</font></td>
			<td width="83%" height="28"><INPUT onclick="javascript:alert('Click on calendar image.')" 
							readOnly size=15  value="<%=enddate%>"name=enddatebox>
							<A onmouseover="window.status='Date Picker';return true;" 
							onmouseout="window.status='';return true;" 
							href="javascript:showCal('EndDate');">
							<IMG height=20 	alt="Select to date." src="images/calendar.gif" width=30 border=0></td>
		</tr>-->
			
          <tr>
				
			<td>
			<INPUT name="save" type="button" value="Show Report" onclick="javascript:fndisplay()"> </td>
			<td>&nbsp;&nbsp;<input type="reset" value="Reset" name=reset ></td>
				
		  </tr>
          
          </table></FORM>  
          <DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%" style="LEFT: 0px; TOP: 178px">
              <TBODY>
              <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE>
</BODY></HTML>
