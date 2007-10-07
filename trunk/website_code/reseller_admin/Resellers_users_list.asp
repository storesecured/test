<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the customers referred by the resellor.
'	Page Name:		   	resellers_users_list.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_sales_history.asp
'	Output Page:	   	resellers_users_list.asp
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
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language=javascript src="include/commonfunctions.js">
	</script>
	
<script language=javascript>
//FUNCTION TO SHOW ALL RECORDS 

function fnShowAll()
{
	
	document.frm.hidpagechanged.value ="True"
	

	    p1 = "0"
	    p2 = "1"		
	    p3 = "1"
		hitlistid = "1"  
document.frm.action = "Resellers_users_list.asp?s=a"+"&p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
document.frm.submit()

}

//Function for paging
function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	
	document.frm.hidpagechanged.value ="True"
	document.frm.SortBy.value=document.frm.hidsortfield.value
		
	document.frm.action = "Resellers_users_list.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	document.frm.submit()
	}
	
//function to show results according to string entered by user in the textbox
function fnShowResult()
{

var textSpeaker 
var ErrMsg,s
	ErrMsg = ""
	document.frm.hidpagechanged.value ="True"
	    p1 = "0"
	    p2 = "1"
	    p3 = "1"
		hitlistid = "1"   
		
// check whitespace
					
	if(isWhitespace(document.frm.txtname.value) == true)
		{	
				ErrMsg=ErrMsg + "Enter some name.\n"
		}
	else 	
		{
			textSpeaker =document.frm.txtname.value
		}	
		
//check for only characters and numeric
	if (isWhitespace(document.frm.txtname.value) == false)
	{
		if( isAlphaNumeric(document.frm.txtname.value) == false)
			{	
				ErrMsg=ErrMsg + "Special characters are not allowed in the search criteria.\n"
			}
	}

	if (ErrMsg != "")
		{
			alert(ErrMsg)
		}
	else
		{
			
			document.frm.action="Resellers_users_list.asp?textval1="+textSpeaker+"&s=s"+"&p1="+ eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
			document.frm.submit()
		}
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
	    theForm.action = "resellers_users_list.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	    theForm.submit( );
	}

	
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">

<%
Dim strid,strcustomers,showresult
showresult="N"

'Retriving the id of the reseller
if trim(Request.QueryString("action"))<>"" then
	strid=trim(Request.QueryString("action"))
end if
if trim(Request.Form("hidid"))<>"" then 
	strid = trim(Request.Form("hidid"))
end if

'Retriving the month & year
compdate1 = trim(Request.Form("fromdate"))
compdate2 = trim(Request.Form("enddate"))
compdate2 = FormatDateTime(compdate2,2)
compdate2 = dateadd ("d",1,compdate2)

if rename="" then
		rename=trim(Request.Form("hidname"))
end if

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
    ordering = "fld_customer_id"
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
		'Response.Write "NOT PAGE ---"&ordering
		'Response.Write "ordering :" & ordering & "<br>"
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


'Retriving customers details from the database depending on the reseller id & date
'sqlid = "select distinct(fld_customer_id) from tbl_reseller_customer_master where "&_
'		" fld_reseller_id="&strid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' and  convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' "&_
'		" order by "&ordering&""
		
'***********************************Code Added  on 22nd Jan*************************************************************		
'sqlid  =  " select distinct(fld_customer_id), fld_Plan_Description from tbl_reseller_customer_master "
'sqlid  =  sqlid  & " where fld_reseller_id="&strid&"  and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' "
'sqlid  =  sqlid  & "  and convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' "
'sqlid  =  sqlid  & " and fld_customer_id in (select store_id from store_settings where  service_type>0 and trial_version=0)"
'sqlid  =  sqlid  & "  order by "&ordering&""

sqlid  =  " select distinct(fld_customer_id), store_user_id , fld_Plan_Description ,sum(fld_amount_pay_to_reseller) as totalsum "
sqlid  = sqlid & "  from tbl_reseller_customer_master a,store_logins b"
sqlid  = sqlid & "  where fld_reseller_id="&strid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' "
sqlid  = sqlid & "  and convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' " 
sqlid  =  sqlid  & " and a.fld_customer_id=b.store_id and a.fld_customer_id in (select store_id from store_settings "
sqlid  =  sqlid  & " where service_type>0 and trial_version=0) "
sqlid  = sqlid & "  group by fld_customer_id, store_user_id, fld_Plan_Description"
sqlid  =  sqlid  & " order by "&ordering&"" 



'Response.Write "<br>sqlid="&sqlid
'response.end
'************************************************************************************
'***********************Code for searching************************

if Request.QueryString("s") = "s" or Request.Form("HidSearch")= "show" then
	Searchval="show"
	showresult="y"
	
	'search the entire word in the list 
	
		scriteria1=trim(Request.QueryString("textval1"))
		scriteria1=fixquotes(scriteria1)
		
		if scriteria1="" then 
			scriteria1=trim(Request.Form("HidScriteria1"))
		end if 
		if scriteria1<> "" then 	
		'-----Commented on 12th March 2005			
'		sqlid = "select distinct(fld_customer_id),fld_Plan_Description from tbl_reseller_customer_master where "&_
'				" fld_reseller_id="&strid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' and  convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"'"&_					
'				" and fld_customer_id LIKE ('%"&scriteria1&"%')  order by "&ordering&""	


		sqlid  =  " select distinct(fld_customer_id), store_user_id , fld_Plan_Description ,sum(fld_amount_pay_to_reseller) as totalsum "
		sqlid  = sqlid & "  from tbl_reseller_customer_master a,store_logins b"
		sqlid  = sqlid & "  where fld_reseller_id="&strid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' "
		sqlid  = sqlid & "  and convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' " 
		sqlid  =  sqlid  & " and a.fld_customer_id=b.store_id and a.fld_customer_id in (select store_id from store_settings "
		sqlid  =  sqlid  & " where service_type>0 and trial_version=0) and fld_customer_id LIKE ('%"&scriteria1&"%') "
		sqlid  = sqlid & "  group by fld_customer_id, store_user_id, fld_Plan_Description"
		sqlid  =  sqlid  & " order by "&ordering&"" 		
'		response.write "sqlid ="& sqlid 
		end if 
		
end if	

' if the user click on the SHOW all button 
if Request.QueryString("s") = "a" then
	 Searchval=""
	 showresult=""  
'	 	sqlid = "select distinct(fld_customer_id) from tbl_reseller_customer_master "&_
'	 			" where fld_reseller_id="&strid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"'"&_
'	 			" and  convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' order by "&ordering&" "

		sqlid  =  " select distinct(fld_customer_id), store_user_id , fld_Plan_Description ,sum(fld_amount_pay_to_reseller) as totalsum "
		sqlid  = sqlid & "  from tbl_reseller_customer_master a,store_logins b"
		sqlid  = sqlid & "  where fld_reseller_id="&strid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' "
		sqlid  = sqlid & "  and convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' " 
		sqlid  =  sqlid  & " and a.fld_customer_id=b.store_id and a.fld_customer_id in (select store_id from store_settings "
		sqlid  =  sqlid  & " where service_type>0 and trial_version=0) "
		sqlid  = sqlid & "  group by fld_customer_id, store_user_id, fld_Plan_Description"
		sqlid  =  sqlid  & " order by "&ordering&"" 

end if
	
'Retriving the reseller name from database depending upon his id
strname = "select fld_first_name+' '+ fld_last_name as name from tbl_reseller_master where fld_reseller_id="&strid&""
'response.write"strname ="&strname 

set rsname = conn.execute (strname)

if not rsname.eof then
	rename=trim(checkencode(rsname("name")))
end if
rsname.close
rsname=nothing

'executing sql query
set rsID=server.CreateObject("ADODB.RecordSet")							
rsID.CursorLocation = 3
rsID.open sqlid,conn,2,2


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
          <form name="frm" method="post" action="">

            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>        
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1" id=TABLE1>
          <tr>
                    <td width="30%" align="middle">

          Store Id</td>
          <td>
          <input type="text" name="txtname">
          </td>
          </tr>
          
          <tr>
          <td></td>
          <td>
          <input type="button" name="show" value="Show" onclick="javascript:fnShowResult()">
     <%if showresult="y" then %>
        <input type="button" name="showall" value="Show All" onclick="javascript:fnShowAll()">
     <% end if%>
          </td>
          </tr>          
          <tr>
          
          Customers Referred by <%= rename%>
		<hr noshade>	
          </tr>
          
          <tr bgColor=#dddddd>
                    <td width="30%" align="middle">
					<strong><%=prefixes(1)%><a href="javascript:goSort('fld_customer_id',1);">Store ID</a></strong>
         </td>
			<td width="30%" align="middle">
				<strong><%=prefixes(2)%><a href="javascript:goSort('store_user_id',2);">Referred Customers</a></strong>
			</td>
			<td width="15%" align="middle">
				<strong><%=prefixes(3)%><a href="javascript:goSort('fld_Plan_Description',3);">Plan</a></strong>
			</td>
			<td width="15%" align="middle">
				<strong><%=prefixes(4)%><a href="javascript:goSort('sum(fld_amount_pay_to_reseller)',4);">Payable Amount</a></strong>
			</td>
			
			<%
			if not rsId.eof then
			
        		' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
												
												rsId.pagesize =50
												p1=Request.QueryString("p1") 
												if p1="" or isnull(p1) then
													p1=0
												end if 	
												
												if  pageno="" or isnull(pageno)  then
													pageno = 1
												elseif  pageno<=0 then
													pageno = 1 
												end if
												if cint(rsId.pagecount) < cint(pageno) then
														pageno = rsId.pagecount
												end if
												dim iNum
												rsId.absolutepage= cint(pageno)
												iNum = pageNo-1
												ctrNum = (iNum*rsId.PageSize)+1
												
												curpage=0
	                 							' ---------
					do while curpage < rsId.PageSize
					intcustomerid  = trim(rsID("fld_customer_id"))
					'************************************************Code Starts******************************************************
					'Commented on 12th March
					'Retriving amount owned by the ESc to reseller for a particular customer 
				'	strGetAmount =	" select sum(fld_amount_pay_to_reseller) as totalsum from tbl_reseller_customer_master "&_
				'					"  where fld_customer_id="&intcustomerid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' "&_
				'					"  and  convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' "
				'	response.write"<br>strGetAmount="&strGetAmount
					
				'	set rsGetAmount=conn.execute(strGetAmount)
					
					
				'	amount=trim(rsGetAmount("totalsum"))			
					amount=trim(rsId("totalsum"))	
					test = split(amount,".")
					if test(1) = "" then
						amount=amount&".00"
					end if
					'************************************************Code Ends******************************************************	
					
					
				'	strGetData = " select fld_customer_id,fld_Plan_Description,fld_plan_user_reseller_id from tbl_reseller_customer_master"&_
				'				"  where fld_customer_id="&intcustomerid&" and convert(datetime,fld_Plan_Transaction_Date,101)>='"& compdate1&"' and "&_
				'				" convert(datetime,fld_Plan_Transaction_Date,101) < '"& compdate2&"' order by fld_reseller_customer_id desc "
				'	set rsGetData=conn.execute(strGetData)
				'	Response.Write "strGetData"&strGetData
					
					'*****************************************************************
					'code ends here
					
					'*****************************************************************
				
				'	if not rsGetData.eof then 
				'		strcustomers=trim(rsGetData("fld_customer_id"))
						strcustomers=trim(rsID("fld_customer_id"))
						planname ="BASIC"	
					'	desc  = trim(rsGetData("fld_Plan_Description")) 
						desc		= trim(rsID("fld_Plan_Description"))
						if desc <> "" then 			
					'		planname = ucase(trim(rsGetData("fld_Plan_Description")))
							planname = ucase(trim(rsID("fld_Plan_Description")))
						end if	
				
					'Retriving the plan name from plan name id			
						'strplanid="select fld_plan_name from tbl_plan_name_master where fld_plan_name_id="&planid&""
						'set rsplanid=conn.execute(strplanid)
						'	if not rsplanid.eof then
						''		planname=trim(rsplanid("fld_plan_name"))
					'		end if	
					'rsplanid.close
					'rsplanid=nothing
'Commented on 12th March
					'Retriving the customer name from customer id
'					strname1="select store_user_id as name from store_logins where Store_id="&strcustomers&" "
'					response.write"<br>strname1="&strname1
'					set rsname1=conn.execute(strname1)
					
'						if not rsname1.eof then
						customername=trim(rsID("store_user_id"))
'						end if	
'						rsname1.close
'						rsname1=nothing
			 %>
          </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
          
          <tr>
          <td width="30%" align="middle">
				<%= strcustomers%>
			</td>
		
          <td width="30%" align="middle">
				<%= customername%>
			</td>
			<td  align="middle">
			<%= planname%>
			</td>
		
			
			<td  align="middle">
				$<%= formatnumber(amount,2)%>
			</td>
			
          </tr>
           <td width="30%" align="middle">
		
		</td>
			<%
				ctrNum = ctrNum + 1
				curpage=curpage+1	
			if not rsID.EOF then
				rsID.movenext	
			end if	
			if rsID.EOF then
				exit do
			end if

			        loop	
				%>			
				<%
			else %>
			 <tr>
			 <td></tD>
			<td ><font color="red" face="arial"> <br>No Record Found</font>
			</td>
			<td></td>
          </tr> 
		<%	
		end if
		
			%>
			
			<% 'rsGetData.close
			'rsGetData=nothing%>
		<tr>
		  		<td>
					<input type="hidden" value="<% =rename%>" name="hidname">
					<input type="hidden" value="<% =strid%>" name="hidid">
					    <input type="hidden" value="<% =compdate1%>" name="fromDate">
						<input type="hidden" value="<% =compdate2%>" name="EndDate" >
						<INPUT Type="hidden" Name="hidSearch" Value="<%=Searchval%>">
						<input type="hidden" Name="hidScriteria1" value="<%=scriteria1%>">
		<!--<br>SortBy--><INPUT Type="hidden" Name="SortBy">
		<!--<br>SortColumn--><INPUT Type="hidden" Name="SortColumn">
		<!--<br>PriorSortBy--><INPUT Type="hidden" Name="PriorSortBy" Value="<%=priorOrder%>">
		<!--<br>hidsortfield--><INPUT Type="hidden" Name="hidsortfield" value="<%=sortfield%>"> <!--second page and n th page navigation-->
		<!--<br>hidsortcol--><INPUT Type="hidden" Name="hidsortcol" value="<%=sortcol%>"> <!--second page and n th page navigation-->
		<!--<br>hidsortorder--><INPUT Type="hidden" Name="hidsortorder" Value="<%=sortorder%>"> <!--second page and n th page navigation-->
		<!--<br>hidpagechanged--><INPUT Type="hidden" Name="hidpagechanged" Value="">

					
				</td>
		
          	<tr>
           <tr>
			<td >&nbsp;
			</td>
          </tr> 
        <TBODY>
               <tr>
		<td align="left">
					 &nbsp;&nbsp;<INPUT name="Back" type="button" value="Back" onclick="javascript:history.back()">
					 
				</td>	
				
		  </tr>
		  
         <tr>
				<td></td>	
				<td >  </td>
				<td> </td>
				<td align="Right" valign='Right'><font face=arial size=2>&nbsp;</font></td>			    
		  </tr>
		  		
           
          </table></FORM>  
          <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
		

			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsID.PageCount,10 %></font></td>
		
          </tr> 
          <% rsID.Close
          rsID=nothing%>
               <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
              
  			</table>


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
