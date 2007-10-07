<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441
server.scripttimeout=1


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page displays sales history of the resellor.
'	Page Name:		   	resellers_reports.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_sales_history.asp
'	Output Page:	    resellers_reports.asp
'	Date & Time:		13 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 

msel7=3
msel8=2

startdate = trim(Request.Form("Start_Date"))
enddate = trim(Request.Form("End_Date"))
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

end if


%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
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
	var maxYear=2020;
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
function fnshow(val)
{
	var ErrMsg
	ErrMsg="";
// new code for date validation
	
if (val==1)
{		
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
					document.frm.action="resellers_show_reports.asp";
					document.frm.submit();
			}
			
}
	if (val==2)
	{
		if(document.frm.frommm.value==0)
			{
			ErrMsg=ErrMsg + "From month is mandatory.\n";
			}
				
		if(document.frm.fromyy.value==0)
			{
			ErrMsg=ErrMsg + "From year is mandatory.";
			}
		if(ErrMsg !="")
			{
			alert(ErrMsg);
			}
		else
			{
			document.frm.action="update_payment.asp";
			document.frm.submit();
			}
	}
}

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
End_Date = FormatDateTime(now(),2)
Start_Date = dateadd("m",-1,End_Date)

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
         <form method=post name="frm" >

        <!-- CURRENT DATES -->
			<input type=hidden name="hidCurrDay" value="<%=day(date)%>">
			<input type=hidden name="hidCurrMonth" value="<%=month(date)%>">				
			<input type=hidden name="hidCurrYear" value="<%=year(date)%>">
			<input type="hidden" name="hidSubmit">
		<!-- CURRENT DATES -->
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>     
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1">
          <tr>
         <font face="arial" size=2> Reseller Sales History</font>
		<hr noshade>	
          </tr>
          
          <!-- code for adding calendar control -->
	<SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
			<tr bgcolor="#ffffff"> 
				<td width="18%" height="28" class="inputname"><font face="Arial, Helvetica, sans-serif" size="2">From Date&nbsp;</font></td>
					<td width="83%" height="28">
					<input name="Start_Date" size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
					<A HREF="#" onClick="cal1.select(document.forms[0].Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" TITLE="cal1.select(document.forms[0].Start_Date,'anchor2','M/d/yyyy',(document.forms[0].Start_Date.value=='')?document.forms[0].Start_Date.value:null); return false;" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
				</td>
			<tr bgcolor="#ffffff"> 
				<td width="18%" height="28" class="inputname"><font face="Arial, Helvetica, sans-serif" size="2">To Date&nbsp;</font></td>
					<td width="83%" height="28"><input name="End_Date" size="10" maxlength=10 value="<%= FormatDateTime(End_Date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
					<A HREF="#" onClick="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" TITLE="cal1.select(document.forms[0].End_Date,'anchor2','M/d/yyyy',(document.forms[0].End_Date.value=='')?document.forms[0].End_Date.value:null); return false;" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>			
				</td>
			</tr>		
			
			<tr>
				<td>
					<br>
					<INPUT name="save" type="button" value="Show Sales" onclick="javascript:fnshow(1)">&nbsp;
					 <input type="reset" value="Reset" name=reset >
				</td>
		  </tr>
		  </table>
		   
		   <TABLE border=1 cellPadding=0 cellSpacing=0 width="100%">
		   
		   <tr>
		   <br>
			<font face="arial" size=2> Reseller Payment Update</font>
			<hr noshade>	
			</tr>
		    
			 <tr>
          <TD class="inputname"><B>Select Month and Year</B></TD>
    <TD class="inputvalue"><select name="frommm" size="1">
       <option selected value="0" >Select Month</option>
       
       <option value="1">January</option>
       <option value="2">February</option>
       <option value="3">March</option>
       <option value="4">April</option>
       <option value="5">May</option>
       <option value="6">June</option>
       <option value="7">July</option>
       <option value="8">August</option>
       <option value="9">September</option>
       <option value="10">October</option>
       <option value="11">November</option>
       <option value="12">December</option>
       </select>
       
        
       <select name="fromyy" size="1">
       <option selected value="0">Select Year</option>
       <option value=2005>2005</option>
       <option value=2006>2006</option>
        <option value=2007>2007</option>


    </select>
       
   
          <tr>		
	
			
          <tr>
				<td>	
					<br>
				<INPUT name="save" type="button" value="Show Payments " onclick="javascript:fnshow(2)">&nbsp;
					 <input type="reset" value="Reset" name=reset >
				</td>
		  </tr>
          </tr>
          </table>
        
		  
		  </table>
          
          </FORM>  
          <DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%" style="LEFT: 0px; TOP: 178px">
              <tbody>
              <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE>
</BODY></HTML>
          <SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("End_Date","date","Please enter a valid end date.");

</script>

