<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the customers referred by the resellor.
'	Page Name:		   	Update_Refund.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_sales_history.asp
'	Output Page:	   	Update_Refund.asp
'	Date & Time:		13 Aug - 14 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
msel15=3
msel16=2
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language=javascript src="../include/commonfunctions.js">
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
document.frm.action = "Update_Refund.asp?s=a"+"&p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
document.frm.submit()

}

//Function for paging
function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	
	document.frm.hidpagechanged.value ="True"
	document.frm.SortBy.value=document.frm.hidsortfield.value
		
	document.frm.action = "Update_Refund.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
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
			
			document.frm.action="Update_Refund.asp?textval1="+textSpeaker+"&s=s"+"&p1="+ eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
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
	    theForm.action = "Update_Refund.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)
	    theForm.submit( );
	}

	function fnChecked(val,obj)
{
			
			var tcount 
			tcount = document.frm.tcount.value;
				
			if (obj.checked == true)
			{
					tcount = tcount + val+",";
			}
			else
			{
					tarray = new Array(100)
					var i
					var tempval
					tempval=""
					delval = val;
						
					tarray = tcount.split(",")
					for(i=0;i<tarray.length-1;i++)
					{
						if(eval((tarray[i])) != eval((delval)))
						{
								tempval = tempval+tarray[i]+",";
						}
							
					}
					
				tcount = tempval;
			}	
			window.document.frm.tcount.value = tcount;
}
function fnRefund()
{

		p1 = "0"
			p2 = "1"
			p3 = "1"
			hitlistid = "1"    
			var tcount
			tcount = window.document.frm.tcount.value;
		
			if (tcount == '')
			{
					alert("No records selected.")
			}
			else
			{
						var ans = confirm("Are you sure to refund the amount?")
						if (ans)
						{
								document.frm.SortBy.value=""
								document.frm.SortColumn.value=""
								document.frm.PriorSortBy.value=""
								document.frm.hidsortfield.value=""
								document.frm.hidsortcol.value=""
								document.frm.hidsortorder.value=""
								document.frm.hidpagechanged.value=""
								document.frm.action = "update_refund.asp?action=action&tcount="+tcount
								document.frm.submit()	
				
						}
			}

}
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">

<%
'code here gets executed when the admin selects to refund the amount and the line item gets deleted from the sys_payments
'********************************Starts here *************************
if trim(Request.QueryString("action"))="action" then 
	tcount = trim(Request.QueryString("tcount"))
	tcount = mid(tcount,1,len(tcount)-1)
	
	'code here to upadate the refund status of the customerids
	sqlUpdate = "Update tbl_request set fld_refund_status=1 where fld_customer_id in ("&tcount&")"
	conn.execute(sqlUpdate)
	
	
	if tcount<> "" then 
		storeId = split(tcount,",")
	'code here to deleting that entry from the sys_payments irrespective of the type of payment made(i.e. customer or normal)
	'********************************Starts here *************************
		for i=0 to ubound(storeId)
			sql_select = "delete from sys_payments where payment_id in (select top 1 sys_payments.Payment_ID from sys_payments inner join sys_billing on sys_payments.store_id = sys_billing.store_id where sys_payments.store_id='"&storeId(i)&"' order by transdate desc)"
			conn.execute(sql_select)
				'code here to update the reseller amount here 
				'***********************************Starts Here ***************************
				sql_reseller = "delete from tbl_reseller_Customer_master where fld_reseller_customer_id in (select top 1 fld_reseller_customer_id from tbl_reseller_customer_master where fld_customer_id='"&storeId(i)&"' order by fld_Plan_Transaction_Date desc)"
				
				
				conn.execute (sql_reseller)	
				'***********************************Ends Here***************************
		Next	
	'********************************ENDS here *************************
	
	end if
	
	
	intflag=1
	
end if
'********************************ENDS here *************************


dim strid,strcustomers,showresult
showresult="N"



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
sqlid = "select fld_customer_id,fld_amount from tbl_request where fld_refund_status=0 "&_
		" order by "&ordering&""
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
					
		sqlid = "select fld_customer_id,fld_amount from tbl_request "&_
				" where fld_customer_id LIKE ('%"&scriteria1&"%') and fld_refund_status=0 order by "&ordering&""
		end if 
		
end if	

' if the user click on the SHOW all button 
if Request.QueryString("s") = "a" then
	 Searchval=""
	 showresult=""  
	 	sqlid = "select fld_customer_id,fld_Amount from tbl_request where fld_refund_status=0  "&_
	 			" order by "&ordering&" "
end if
	

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
          <TD class=pagetitle height=400 vAlign=top width=100%>
          <table border="1" id=TABLE1 width=100%>
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
          
          Refund Request
		<hr noshade>	
          </tr>
          
          
          <tr bgColor=#dddddd>
                    <td align="middle">
					<strong><%=prefixes(1)%><a href="javascript:goSort('fld_customer_id',1);">Store ID</strong></a>
         </td>
			<td align="middle">
				<%=prefixes(2)%><a href="javascript:goSort('fld_amount',2);">Amount</a> 
			</td>
			<td align="middle">
				Refund
			</td>
		</tr>
		<%if rsId.eof then%>
			 <tr>
			 <td></tD>
			<td ><font color="red" face="arial"> <br>No Record Found</font>
			</td>
			<td></td>
          </tr> 
          <%
          End if%>	
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
					do while curpage<rsId.PageSize
					intcustomerid  = trim(rsID("fld_customer_id"))
					amount =  trim(rsID("fld_amount"))
					%>
          
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
          
                 
          
          <tr>
          
          <td align="middle">
				<%= intcustomerid%>
			</td>
		
          	<td  align="middle">
				<%if amount<>"" then 				
					Response.Write "$ "&FormatNumber( amount,2)
				end if%>	
					
			</td>
		
			 <td align="middle">
<%
			if trim(tcount) = "" then 
			'Response.Write "Status"&rsid("fld_customer_id")&"-"&Status
			
					Response.Write "&nbsp;<input type=checkbox class=INPUT_check border=0 "&_
					" name='Id'  value='"&rsid("fld_customer_id")&"' "&_
					" onclick=javascript:fnChecked('"&rsid("fld_customer_id")&"',this)>"
			else		
					dim arrayofcount
					arrayofcount = split(tcount,",")

					for i = 0 to ubound(arrayofcount)-1
						if (cint(arrayofcount(i)))=	cint((rsid("fld_customer_id"))) then
								status = "Checked"
					exit for
						else
								status = ""		
						end if
								next 
											
						Response.Write "&nbsp;<input type=checkbox class=INPUT_check border=0 "&_
						" name='Id' "&status&" value='"&rsid("fld_customer_id")&"' "&_
						" onclick=javascript:fnChecked('"&rsid("fld_customer_id")&"',this)>"
			end if			
%>
         
        
		
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
			<!-- </tr>
			 <tr>
			 <td></tD>
			<td ><font color="red" face="arial"> <br>No Record Found</font>
			</td>
			<td></td>
          </tr> -->
		<%	
		end if
		
			%>
			
			
		<tr>
		  		<td>
					<INPUT Type="hidden" Name="tcount" Value="<%=tcount%>">
					
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
					 &nbsp;&nbsp;<INPUT name="Refund" type="button" value="Refund" onclick="javascript:fnRefund()">
					&nbsp;&nbsp;<INPUT name="Reset" type="Reset" value="Reset" >
				</td>	
				
		  </tr>
		  
         <tr>
				<td></td>	
				<td >  </td>
				<td> </td>
				<td align="Right" valign='Right'><font face=arial size=2></font></td>			    
		  </tr>
		  		
           
          </table></FORM>  
          <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
		

			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsID.PageCount,50 %></font></td>
		
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
<%if intflag=1 then%>
	<script language ="">
		alert("The refund status has been updated")
		document.location.href = "update_refund.asp"
		
		
	</script>
<%end if%>