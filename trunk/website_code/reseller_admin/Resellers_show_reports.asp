<!--#include file = "include/ESCheader.asp"-->


<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the amount payable ,amount due to the resellor.
'	Page Name:		   	resellers_show_reports.asp
'	Version Information:EasystoreCreator
'	Input Page:		    resellers_reports.asp
'	Output Page:	    resellers_show_reports.asp
'	Date & Time:		13 Aug - 14 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
msel7=3
msel8=2
%>
<HTML>
	<HEAD>
		<TITLE>Easystorecreator</TITLE>
		<META content=no-cache name=Pragma>
		<META content=no-cache http-equiv=pragma>
		<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
		<LINK href="images/style.css" rel=stylesheet type=text/css>
		<SCRIPT language=JavaScript src="images/script.js" type=text/javascript></SCRIPT>
		<script language ="javascript">
			function fnShow(id)
			{
			document.frm.action = "resellers_users_list.asp?action="+id
			document.frm.submit()
			}

			function jsHitListAction(p1,p2,p3,hitlistid)
				{
				var str ="";
				document.frm.hidpagechanged.value ="True"
				document.frm.action = "resellers_show_reports.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
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
						theForm.SortColumn.value = columnNum;
						theForm.action = "resellers_show_reports.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
						theForm.submit( );
				}				
			</script>
			<META content="MSHTML 5.00.3813.800" name=GENERATOR>
	</HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
Dim sqlGetData,strfirst,frommonth,fromyear,tomonth,toyear,intresellerid

'Retriving the values for from year to to year from resellers reports.asp
fromDate = trim(Request.Form("start_date"))
EndDate = trim(Request.Form("End_Date"))

'ret from hidden variable
if fromDate = "" then 
	fromDate = trim(Request.Form("fromDate"))
end if	
if EndDate = "" then 
	EndDate = trim(Request.Form("EndDate"))
end if
compdate1 = fromDate
EndDate = FormatDateTime(EndDate,2)
EndDate = dateadd ("d",1,EndDate)
compdate2 = EndDate

'****************************Code for sorting*********************

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
    ordering = "fld_first_name"
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
		if trim(sortorder)= "ASC" then
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


'Retriving the resellers list from the database
'sqlGetData = "select fld_first_name+' '+fld_last_name as name,fld_reseller_id from tbl_reseller_master where convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' and  convert(datetime,fld_Plan_Transaction_Date,101) <= '"& compdate2&"' order by "&ordering&" "
sqlGetData = "select fld_first_name+' '+fld_last_name as name,fld_reseller_id from tbl_reseller_master order by "&ordering&" "

set rsGetData = server.CreateObject("ADODB.RecordSet")							
rsGetData.CursorLocation = 3

'executing sql query
rsGetData.open sqlGetData,conn,2,2

%>
<DIV id=overDiv style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
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
								<UL><BR><BR></UL>
							</TD>
						</TR>
					</TBODY>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD>
				<!--#include file="incESCmenu.asp"-->
		<TR>
	    <TD>
		    <TABLE border=0 cellPadding=0 cellSpacing=0>
			    <TBODY>
				  </TBODY>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD>
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TBODY>
						<TR vAlign=top>
							<TD rowSpan=2 width=180>
			          <form name="frm" method="post">
					      <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
							    <TBODY>        
                  </TBODY>
								</TABLE>
							</TD>
		          <TD height=15 vAlign=top width=570></TD>
						</TR>
         <!--sending month & year to another page -->
        <!--<input type="hidden" value="<% =frommonth%>" name="frommonth" >-->
        <input type="hidden" value="<% =compdate1%>" name="fromDate">
        <input type="hidden" value="<% =compdate2%>" name="EndDate" >
        <!--<input type="hidden" value="<% =fromyear%>" name="fromyear" >
        
        <input type="hidden" value="<% =tomonth%>" name="tomonth" >
        <input type="hidden" value="<% =toyear%>" name="toyear" >-->
        
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
						<table border="1" id=TABLE1 width=100%>
							<tr>
								Reseller Payout - <%= compdate1%> to <%= compdate2%>
								<hr noshade>	
							</tr>         
							<tr bgColor=#dddddd>
								<td width="30%" align="middle">
									<strong><%=prefixes(1)%><a href="javascript:goSort('fld_first_name',1);">Resellers</strong></a>
								</td>		
								<td  align="middle">
									Customers Referred(Paid/NonPaid)
								</td>			
								<td align="middle">
									Net Payable Amount
								</td>
						</tr>
						<tr>
							<td >&nbsp;</td>
						</tr> 
          <%if not rsGetData.eof then	        
          ' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
												
												rsGetData.pagesize =50
												p1 = Request.QueryString("p1") 
												if p1 = "" or isnull(p1) then
													p1 = 0
												end if 	
												
												if  pageno = "" or isnull(pageno)  then
													pageno = 1
												elseif  pageno <= 0 then
													pageno = 1 
												end if
												
												if cint(rsGetData.pagecount) < cint(pageno) then
														pageno = rsGetData.pagecount
												end if
												dim iNum
												rsGetData.absolutepage = cint(pageno)
												iNum = pageNo-1
												ctrNum = (iNum*rsGetData.PageSize) + 1
												
												
												curpage = 0
	                 							' ---------
					do while curpage<rsGetData.PageSize
						strfirst = trim(rsGetData("name"))
						intresellerid = trim(rsGetData("fld_reseller_id"))
		  %>
		   
         <tr>
          <td align="middle">
						<a href="javascript:fnShow(<%= intresellerid%>)" title="Click to view customers list"><%= strfirst%><a>
					</td>			
				<% 
				'Retriving the count of customers referrd by the reseller
				Dim  intTotCnt, intNotPaid
				Dim  rstotcust

				strtotcustomers = "select count(distinct(fld_customer_id))as count from tbl_reseller_customer_master"&_
								" where fld_reseller_id="&intresellerid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' and  convert(datetime,fld_Plan_Transaction_Date,101) <= '"& compdate2&"' "
				
				'strtotcustomers  =  "select  IsNull(count(store_id) ,0)  as  count  from store_settings where reseller_id = "&intresellerid&" "				
				'response.write"strtotcustomers ="&strtotcustomers 

				Set  rstotcust  =  conn.execute(strtotcustomers)
				If Not rstotcust.Eof  Then
					intTotCnt = Trim(rstotcust("count"))
				End if	
				'***************************Code  Added  on  22nd Jan
				strcustomers  = "select count( distinct ( fld_customer_id ) )  as count  from tbl_reseller_customer_master "
				strcustomers  = strcustomers  & " where fld_reseller_id="& intresellerid &"  and convert(datetime, "
				strcustomers  = strcustomers  & "  fld_Plan_Transaction_Date,101)>='"& compdate1&"' and convert(datetime,fld_Plan_Transaction_Date,101) <=  '"& compdate2&"' "
				strcustomers  = strcustomers  & "  and fld_customer_id in (select store_id from store_settings where  service_type>0 and trial_version=0)"			

				'**************************************************************************************************************	
				'response.write"strcustomers="&strcustomers	
				set rscustomers = conn.execute(strcustomers)
				if not rscustomers.eof then
					recount = trim(rscustomers("count"))
					intNotPaid =  intTotCnt  -  recount
				%>
				<% rscustomers.close
				rscustomers = nothing%>
			<td  align = "middle">
				<%= recount%>/<%=intNotPaid %>
			</td>
			<% end if%>
			<%
			'Retriving the net payable amount of the reseller

			stramount = "select sum(fld_amount_pay_to_reseller)as sum from tbl_reseller_customer_master where convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' and  convert(datetime,fld_Plan_Transaction_Date,101) <= '"& compdate2&"' and fld_reseller_id="&intresellerid&""
			
			'Response.Write "stramount"&	stramount
			
			set rsamount = conn.execute(stramount)
				amount = "0.00"
			
				
				if not rsamount.eof then
						if isNull(trim(rsamount("sum"))) = False then 
							amount = trim(rsamount("sum"))
							test = split(amount,".")
							if test(1) = "" then
								amount = amount&".00"	
							end if
						end if
			%>
			<% rsamount.close
			rsamount = nothing%>
			<td align="middle">
				$ <%= formatnumber(amount,2)%> 
			</td>
			<%
				end if
			%>
     </tr>
				<td  align="middle"></td>
		 <tr>
			<td  align="middle"></td>      
			<td  align="middle"></td>			
     </tr>          
     <tr>
			<td  align="middle"></td>			
			<td  align="middle"></td>			
     </tr>          
     <tr>
			<td  align="middle"></td>			
			<td  align="middle"></td>			
     </tr>
     <tr>
			<td  align="middle"></td>					
			<td>
			<%	
				
				ctrNum = ctrNum + 1
				curpage = curpage+1	
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
		<td >&nbsp;</td>
  </tr>         
	<tr>
		<td><input type="button" name="back" value="back" onclick="javascript:history.back()"></td>	
		<td>  </td>
		<td> </td>
		<td align="Right" valign='Right'><font face=arial size=2>&nbsp;</font></td>			    
	</tr>
  <tr>
		<td><br><input type="hidden" value="<%= frommonth%>" name="hidmm"></td>	
	   <tr>         
	     <!--<br>SortBy--><INPUT Type="hidden" Name="SortBy">
			<!--<br>SortColumn--><INPUT Type="hidden" Name="SortColumn">
			<!--<br>PriorSortBy--><INPUT Type="hidden" Name="PriorSortBy" Value="<%=priorOrder%>">
			<!--<br>hidsortfield--><INPUT Type="hidden" Name="hidsortfield" value="<%=sortfield%>"> <!--second page and n th page navigation-->
			<!--<br>hidsortcol--><INPUT Type="hidden" Name="hidsortcol" value="<%=sortcol%>"> <!--second page and n th page navigation-->
			<!--<br>hidsortorder--><INPUT Type="hidden" Name="hidsortorder" Value="<%=sortorder%>"> <!--second page and n th page navigation-->
			<!--<br>hidpagechanged--><INPUT Type="hidden" Name="hidpagechanged" Value="">
       </table>
		</FORM>            
		<TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
		 <tr>        
				<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsGetData.PageCount,10 %></font></td>
				</tr> 
        <% rsGetData.Close
        rsGetData=nothing%>
        <TBODY>
         <TR>
					<TD height=20></TD>
					</TR>
				</TBODY>
			</TABLE>
		</TD>
	</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TABLE>
</BODY>
</HTML>

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

