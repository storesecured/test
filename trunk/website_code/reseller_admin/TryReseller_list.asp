<!--#include file = "../include/ESCheader.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the list of all the registered resellers.
'	Page Name:		   	reseller_list.asp
'	Version Information:EasystoreCreator
'	Input Page:		    escadmin.asp
'	Output Page:	    reseller_list.asp
'	Date & Time:		10 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK href="images/style.css" rel=stylesheet type=text/css>
<META content="MSHTML 5.00.3813.800" name=GENERATOR>
<script language="javascript">
function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	document.frm.action = "reseller_list.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	document.frm.submit()
	}

</SCRIPT>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">


<%

dim sqlGetData,strfirst,strwebsite,strmail,sqlmode,intmode,intresellerid,rsgetdata

'Retriving the reseller information from the database
sqlGetData = "select fld_first_name+' '+fld_Last_Name as name,fld_website,fld_mail,fld_reseller_id from TBL_Reseller_Master"
'set rsGetData = conn.execute(sqlGetData)
'Response.Write "sqlGetData"&sqlGetData

set rsGetData=server.CreateObject("ADODB.RecordSet")							
rsGetData.CursorLocation = 3

'executing sql query
rsGetData.open sqlGetData,conn,2,2

%>
<DIV id=overDiv 
style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  <TBODY>
  <TR>
    <TD class=title>
    <form name="frm" method="post" action="" >
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
          <TD class=title><B>Easystorecreator</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%">
            <UL><BR><BR></UL></TD></TR></TBODY></TABLE></TD></TR>
            
  <TR>
    <TD>
     <!--#include file="incESCmenu.asp"--></TD></TR>
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
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1" width="100%" > 
          <tr>
          Resellers List 
		<hr noshade>	
          </tr>
          
          <tr bgColor=#dddddd>
			<td >
				Reseller Name
			</td>
			<td >
				Buisness Name
			</td>
			<td>
				Email
			</td>
			<td>
				Payment Mode
			</td>
			
			<td>
				
			</td>
          </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
	<%
				if not rsGetData.eof then
				' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
												'Response.Write "pageno"&pageno
												'Response.Write "pageno"&rsgetdata.PageSize
												
												rsGetData.pagesize =2
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
						 do while curpage<rsGetData.PageSize
							strfirst=trim(rsGetData("name"))
							strwebsite=trim(rsGetData("fld_website"))
							strmail=trim(rsGetData("fld_mail"))
							intresellerid = trim(rsGetData("fld_reseller_id"))
							sqlmode = "select fld_payment_mode from TBL_ESC_Reseller_payment_mode where fld_reseller_id="&intresellerid&""
							set rsmode=conn.execute(sqlmode)
							ModeName ="Not Set"	
							if not rsmode.eof then
								intmode = trim(rsmode("fld_payment_mode"))
								if intmode ="2" then 
									ModeName ="By Check"	
								elseif intmode = "1" then 
									ModeName ="By Paypal"	
								end if
							end if
			%>
          <tr>
			
			<td>
			<%=strfirst%> 		
			</td>
			<td >
			<%=strwebsite%>
			</td>
			<td>
			<%=strmail%>
			</td>
			
			<td>
			<%if ModeName ="Not Set" then
				 Response.Write Modename
			end if	
			if ModeName ="By Check" then%>
			<a href="reseller_payment_check.asp?action=<%=intresellerid%>" title="Click here for billing info"><%= ModeName%></a>
			<%end if
			if ModeName ="By Paypal" then%>
			<a href="reseller_payment_paypal.asp?action=<%=intresellerid%>" title="Click here for billing info"><%= ModeName%></a>
			<%end if%>
			
			</td>
			<td>
			<A href="edit_reseller.asp?action=<%=rsGetData("fld_reseller_id")%>" title="Click to edit reseller">Edit</a>
			/<A href="show_month.asp?action=<%=rsGetData("fld_reseller_id")%>" title="Click to update Payment Status">Update Payment</a>
			</td>
			
          </tr>
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
			
            <tr>
            <td>
            <td>
			<td ><!--<font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsGetData.PageCount,5 %></font>-->
			
			</td>
			<td><td>
          </tr> 
        
          <tr>
				<td></td>	
				<td >  </td>
				<td> </td>
				<td align="Right" valign='Right'><font face=arial size=2></font></td>			    
		  </tr>
          
          </table>
          </form>
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
        
			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsGetData.PageCount,5 %></font></td>
		
          </tr> 
        
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
	'Response.Write("<br>" & counterstart)
	'Response.Write("<br>" & counterend)
			
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

