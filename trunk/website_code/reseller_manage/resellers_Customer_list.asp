<!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page displays list of the customer for the resellers site
'	Page Name:		   	resellers_customer_list.htm
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_customer_list.htm
'	Date & Time:		5 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel5=3
mSel6=2
%>


<script language="javascript" src="../include/commonfunctions.js"></script>

<script language="javascript">



function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	
	tcount = document.frmcustomerlist.tcount.value;
		document.frmcustomerlist.hidpagechanged.value ="True"
		document.frmcustomerlist.SortBy.value=document.frmcustomerlist.hidsortfield.value
	document.frmcustomerlist.action = "resellers_customer_list.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	document.frmcustomerlist.submit()
	}
	
	//function for sorting 
function goSort( fldName, columnNum )
	{
	    p1 = "0"
	    p2 = "1"
	    p3 = "1"
		hitlistid = "1"    
	    theForm = document.frmcustomerlist;
	    
	    theForm.SortBy.value = fldName;
	    //theForm.SortMsg.value = headerText;
	    theForm.SortColumn.value = columnNum;
	    theForm.action = "resellers_customer_list.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	    theForm.submit( );
	}

</script>

<HTML><HEAD><TITLE>Reseller Admin</TITLE>


<SCRIPT language=JavaScript src="images/CalendarPopup.js"></SCRIPT>

<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%

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
    ordering = "store_id"
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


'code for display customer list.

dim sqlgetdata,sqlgetcustid,intresellerid,rsdisplaydata,intcustid,rsdisplaydataid
intresellerid = trim(session("resellerid"))

sqlgetcustid="select distinct(store_id) from store_settings where reseller_id="&intresellerid&" order by "&ordering&""

set rsdisplaydataid=server.CreateObject("adodb.recordset")

rsdisplaydataid.CursorLocation = 3

'executing sql query
rsdisplaydataid.Open sqlgetcustid,conn,2,2

	

%>
<DIV id=overDiv 
style="Z-INDEX: 1000; VISIBILITY: hidden; POSITION: absolute"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750 id=TABLE1>
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
    <!--#include file="incmenu.asp" -->
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
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>
            
                  
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Customers
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action=""   method=post name="frmcustomerlist"><INPUT name=Form_Name type=hidden>
               <INPUT name=redirect type=hidden value=.asp>
               <INPUT name=records type=hidden value=25> 
              <TBODY>
              <TR>
                <TD colSpan=10>
                  <TABLE>
                    <TBODY>
                    <TR>
                      <TD height=15 vAlign=top width=570></TD></TR></TBODY></TABLE>
   
          <table border="1" id=TABLE2 style="WIDTH: 614px; HEIGHT: 50px">
          <tr bgColor=#dddddd>
			
			<td align="middle"  nowrap class="bodybold"><strong><%=prefixes(1)%><a href="javascript:goSort('fld_customer_id',1);">Store Id</strong></a>
			</td>
			<td>
			Store User Id
			</td>
			<td>
		First Name
			</td>
			<td  align="middle">
				Last Name
			</td>
			<td  align="middle">
				Service
			</td>
			<td  align="middle">
				Email
			</td>
			<td >
			</td>
			
          </tr>
          <%
       if not rsdisplaydataid.eof then
			
							' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
												
												rsdisplaydataid.pagesize =20
												p1=Request.QueryString("p1") 
												if p1="" or isnull(p1) then
													p1=0
												end if 	
												
												if  pageno="" or isnull(pageno)  then
													pageno = 1
												elseif  pageno<=0 then
													pageno = 1 
												end if
												
												if cint(rsdisplaydataid.pagecount) < cint(pageno) then
														pageno = rsdisplaydataid.pagecount
												end if
												dim iNum
												rsdisplaydataid.absolutepage= cint(pageno)
												iNum = pageNo-1
												ctrNum = (iNum*rsdisplaydataid.PageSize)+1
												
												
												curpage=0
                 							'code ends here
                 							
                 							
                 						
	                 					
	do while curpage<rsdisplayDataid.PageSize
				intcustid=rsdisplaydataid("store_id")
				'Retrieve the store user id from store_logins from store_id
				sqlid="select Store_id,Store_User_Id from Store_logins where Store_id="&intcustid&" "
				set rsid= conn.execute(sqlid)
					if not rsid.eof then
						storeuserid=trim(rsid("Store_User_Id"))
						storeid=trim(rsid("store_id"))	
					end if
				First_Name = "-"
				Last_Name = "-"
				Company = "-"
				Email  = "-"
				'code for retriving company name from store_settings table
				sqlgetcompname="select store_company from store_settings left join store_logins"&_ 
						   " on store_settings.store_ID="&intcustid&" and store_logins.store_ID="&intcustid&""
						set rsdata=conn.execute(sqlgetcompname)   
						if not rsdata.eof then
						Company=trim(rsdata("store_company"))
						end if
						
				'code ends here
				
				strGetData = " select fld_Plan_Description from tbl_reseller_customer_master"&_
								"  where fld_customer_id="&intcustid&" order by fld_reseller_customer_id desc "
				set rsGetData=conn.execute(strGetData)
					
					'*****************************************************************
					'code ends here
					
					'*****************************************************************
					planname = "BASIC"
					if not isNULL(fld_Plan_Description) then 
						desc  = trim(rsGetData("fld_Plan_Description"))
						if desc <> "" then 			
							planname = trim(rsGetData("fld_Plan_Description"))			
						end if	
					end if							
						
				sqlgetdata="select store_id,first_name,last_name,email from Sys_Billing where store_id="&intcustid
				set rsdisplaydata=conn.execute(sqlgetdata)
				if not rsdisplaydata.eof then
						First_Name = trim(rsdisplaydata("first_name"))
						Last_Name = trim(rsdisplaydata("last_name"))
						Email  = trim(rsdisplaydata("Email"))
						store_id = trim(rsdisplaydata("store_id"))
				end if	
						
					%> 
					    </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
					 <tr>
					
				 
			      <td align="middle">
			 <%= intcustid%>
			 </td>
			 <td align="middle">
			 <%= storeuserid%>
			 </td>
			 <td  align="middle">
				<%= First_Name%>
			</td>
			<td  align="middle">
				<%=Last_Name %>
			</td>
			<td  align="middle">
				<%=ucase(planname)%>
			</td>
			<!--<td  align="middle">
				<%=Company%>
			</td>-->
			<td  align="middle">
				<%=Email%>
			</td>
			<td >
				<%if trim(store_id)<> "" and planname <> "BASIC" then %>
				<a href="edit_Customer.asp?action=<%=intcustid%>">View</a>
				<%end if%>
			</td>
			
          </tr>
         <%			             
				ctrNum = ctrNum + 1
				curpage=curpage+1	
		         
		    if not rsdisplaydataid.EOF then
				rsdisplaydataid.movenext
			end if
			if rsdisplayDataid.EOF then
				exit do
			end if
				
			  loop
	  else
	  %>
	   <tr>
			<td  align="middle">
			</td>
			<td  align="middle">
			</td>
			<td  align="middle">
			<font color="red">No Record Found</font>
				</td>
			<td  align="middle">
			</td>
			<td >
			</td>
		  </tr>
          
          </table>
          
      <%end if    %>
      
 <!--<br>SortBy-->     <INPUT Type="hidden" Name="SortBy">
<!--<br>SortColumn-->  <INPUT Type="hidden" Name="SortColumn">
<!--<br>PriorSortBy--> <INPUT Type="hidden" Name="PriorSortBy" Value="<%=priorOrder%>">
<!--<br>hidsortfield--><INPUT Type="hidden" Name="hidsortfield" value="<%=sortfield%>"> <!--second page and n th page navigation-->
<!--<br>hidsortcol-->  <INPUT Type="hidden" Name="hidsortcol" value="<%=sortcol%>"> <!--second page and n th page navigation-->
<!--<br>hidsortorder--><INPUT Type="hidden" Name="hidsortorder" Value="<%=sortorder%>"> <!--second page and n th page navigation-->
						<!--<INPUT Type="hidden" Name="HidSearch" Value="<%=Searchval%>">
						<input type="hidden" Name="HidScriteria1" value="<%=scriteria1%>">-->
<!--<br>hidpagechanged--><INPUT Type="hidden" Name="hidpagechanged" Value="">
<INPUT Type="hidden" Name="tcount" Value="<%=tcount%>">
     
      
      
               <tr>
				<td></td>	
				<td >  </td>
				<td> </td>
				<td align="Right" valign='Right'><font face=arial size=2></font></td>			    
			   </tr>
		</table>
		</FORM>
  
<TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
        
			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsdisplaydataid.PageCount,10 %></font></td>
		<% rsdisplaydataid.Close
		rsdisplaydataid=nothing%>
          </tr> 
 <DIV id=testdiv1 style="BACKGROUND-COLOR: white; POSITION: absolute; VISIBILITY: hidden; layer-background-color: white"></DIV>

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
	'if Cint(MaxPages)-Cint(CounterEnd) < cint(DisplayPageCount) then
	'	Response.Write "<font face=arial size=2><a href='javascript:jsHitListAction("& p1 &"," & CounterEnd+1 & "," & p3 &"," & hitlistid &")'><font size=2 face=arial>Next</font></a></font>"
	'	Response.Write "<font > |</font> "			
	'else
		Response.Write "<font face=arial size=2><a href='javascript:jsHitListAction("& p1 &"," & CounterEnd+1 & "," & p3 &"," & hitlistid &")'><font size=2 face=arial >Next</font></a></font>"
		Response.Write "<font face=arial size=2> |</font> "		
		Response.Write "<font face=arial size=2><a href='javascript:jsHitListAction("& p1 &"," & maxpages & "," & p3 &"," & hitlistid &")'><font size=2 face=arial >Last</font></a></font>"
		Response.Write "<font face=arial size=2> |</font> "				
	'	end if		
	end if
end if			
end sub

%>

