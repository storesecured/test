<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page displays sales history of the resellor.
'	Page Name:		   	reseller_sales_history.htm
'	Version Information: EasystoreCreator
'	Input Page:		    resellers_sales_report.asp
'	Output Page:	    resellers_sales_report.asp
'	Date & Time:		11 and 12 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel7=3
mSel8=2

%>
<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
<LINK href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language="javascript">
function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	document.frmsalesreport.action = "resellers_sales_report.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	document.frmsalesreport.submit()
	}

</SCRIPT>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
dim frommonth,fromyr,tomonth,toyr,sqlescdata,sqlstatus,rsescdata,statusid
dim paidamt,balanceamt,strstatname

fromDate = trim(Request.Form("start_date"))
EndDate = trim(Request.Form("End_date"))
if fromDate = "" then 
	fromDate = trim(Request.Form("fromDate"))
end if	
if EndDate = "" then 
	EndDate = trim(Request.Form("EndDate"))
end if
compdate1 = fromDate

compdate2 = EndDate

month1 = month(fromDate)
month2 = month(EndDate)
Year1 = year(fromDate)
year2 = year(EndDate)

'query to retrive data from table tbl_esc_reseller_payment.
intresellerid = session("Resellerid")
			
sqlescdata = " select sum(fld_amount_pay_to_reseller) as amount,month(fld_Plan_Transaction_Date) aS months,"&_
			 " year(fld_Plan_Transaction_Date) as years from tbl_reseller_customer_master WHERE fld_reseller_id='"&intresellerid&"' "&_
		  	 " and month(fld_Plan_Transaction_Date) between '"&month1&"' and '"&month2&"' "&_
			 " and year(fld_Plan_Transaction_Date) between '"&Year1&"' and '"&Year2&"' GROUP BY month( fld_Plan_Transaction_Date),year(fld_Plan_Transaction_Date)"


set rsescdata=server.CreateObject("ADODB.RecordSet")
rsescdata.CursorLocation = 3

'executing sql query
rsescdata.open sqlescdata,conn,2,2
'Response.Write sqlescdata
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
    </td>
    </tr>
    
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
          <TD class=pagetitle height=400 vAlign=top>My sales Payout from <%= compdate1%> to <%= compdate2%>
 
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="100%"><FORM action=""   method=post name="frmsalesreport"><INPUT name=Form_Name type=hidden> <INPUT 
              name=redirect type=hidden value=.asp> <INPUT 
              name=records type=hidden value=25> 
              <TBODY>
              <TR>
                <TD colSpan=10>
                  <TABLE>
                    <TBODY>
                      <TD height=15 vAlign=top width=570></TD></TR>
   
          <table border="1" width=100%>
          <tr bgColor=#dddddd>
			<td >
				Month
			</td>
			<td >
				Total
			</td>
			
			<td >
				Amount Paid
			</td>
			<td>
				Balance Amount 
			</td>
			
			<td >
				Payment Status
			</td>
			
          </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
         <% 
        if not rsescdata.eof then 
        
         ' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
												rsescdata.pagesize = 50
												p1=Request.QueryString("p1") 
												if p1="" or isnull(p1) then
													p1=0
												end if 	
												
												if  pageno="" or isnull(pageno)  then
													pageno = 1
												elseif  pageno<=0 then
													pageno = 1 
												end if
												
												if cint(rsescdata.pagecount) < cint(pageno) then
														pageno = rsescdata.pagecount
												end if
												dim iNum
												rsescdata.absolutepage= cint(pageno)
												iNum = pageNo-1
												ctrNum = (iNum*rsescdata.PageSize)+1
												
												
												curpage=0
	                 							' ---------
				
				do while curpage<rsescdata.PageSize
				profitamt=trim(rsescdata("amount"))
				if profitamt<> "" then 
					test = split(profitamt,".")
					if test(1) = "" then
						profitamt=profitamt&".00"
					end if
				end if	
				strMonth=trim(rsescdata("months"))
				currentmonth =strMonth
				currentyear = trim(rsescdata("years"))
				strMonth = monthname(strMonth)
				
					
'*******************************CODE STARTS HERE***********************************************************************
'code here to reterive the net profit 
strgetdata= " select sum(fld_amount_paid) AS paid FROM TBL_ESC_Reseller_Payment "&_
			" WHERE fld_reseller_id='"&intresellerid&"' and fld_month = '"&currentmonth&"' and fld_year = '"&currentyear&"'"
set rsselect=conn.execute(strgetdata)
if not isNull(rsselect("paid")) then 
	paidamt=trim(rsselect("paid"))
	balanceamt = csng(profitamt)-csng(paidamt)
	test = split(paidamt,".")
	paidamt = formatnumber(paidamt,2)
	if test(1) = "" then
		paidamt=paidamt&".00"
	end if
	test = split(balanceamt,".")
	if test(1) = "" then
		balanceamt	= formatnumber(balanceamt,2)
		balanceamt=balanceamt&".00"
	end if
else

	paidamt = "0.00"
end if

		rsselect.close
		set rsselect=nothing
					if balanceamt="0" then 
						strstatname="Complete"
					else
						strstatname="Pending"
					end if
 
'*******************************CODE ENDS HERE***********************************************************************
			
		%>
          <tr>
          <td width="30%">
				<%=strMonth%>
			</td>
			<td width="30%">
				<%=formatnumber(profitamt,2)%>
			</td>
			<td >
				$ <%=paidamt%>
			</td>
			
			<td >
				$ <%=balanceamt%>
			</td>
			    <td >
				<%=strstatname %>
			<!--<a href="edit_customer.asp" title="Click to edit customer">Edit</a> / <a href="#" title="Click to delete customer">Delete</a>-->
			</td>
		 </tr>
		 <%
		
			ctrNum = ctrNum + 1
				curpage=curpage+1	
			if not rsescdata.EOF then
				rsescdata.movenext	
			end if	
			if rsescdata.EOF then
				exit do
			end if
				loop	
			
			else	
	   	%>	
			<td></td>
			<td></td>
			<td>
			<font color="red">No records found.</font>
			</td>
			
			<% end if%>
          <%
          'rsescdata.close
          'set rsescdata=nothing
          %>
          
           <td width="30%">
		
		</td>
		   
           <tr>
			<td >&nbsp;
			</td>
          </tr> 
           
           <tr>
			<td ><input type="button" name="back" value="Back" onclick="javascript:history.back()">
			</td>
          </tr> 
          <tr>
				<td></td>	
				<td >  </td>
				<td> </td>
				<td align="Right" valign='Right'><font face=arial size=2></font></td>			    
		  </tr>
        </table>
    
            
              <TR>
                <TD class=inputname></TD></TR>
              <TR>
              <input type="hidden" value="<% =compdate1%>" name="fromDate">
        <input type="hidden" value="<% =compdate2%>" name="EndDate" >
                <TD 
  height=20></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></FORM>
  
  <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
		

			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsescdata.PageCount,5%></font></td>
		
          </tr> 
  			</table>

<DIV id=testdiv1 
style="BACKGROUND-COLOR: white; POSITION: absolute; VISIBILITY: hidden; layer-background-color: white"></DIV>


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
	end if
end if			
end sub

%>


