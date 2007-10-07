<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the customers referred by the resellor.
'	Page Name:		   	selectstore.asp
'	Version Information:EasystoreCreator
'	Input Page:		    
'	Output Page:	   	selectstore.asp
'	Date & Time:		
'	Created By:			Devki Anote
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
msel7=3
msel8=2
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
<!--<LINK href="images/style.css" rel=stylesheet type=text/css>-->
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
		tcount = document.frm.tcount.value;
document.frm.action = "selectstore.asp?s=a"+"&p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
document.frm.submit()

}
7
//Function for paging
function jsHitListAction(p1,p2,p3,hitlistid)
	{
	var str ="";
	tcount = document.frm.tcount.value;
	document.frm.hidpagechanged.value ="True"
	document.frm.SortBy.value=document.frm.hidsortfield.value
	document.frm.action = "selectstore.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
	document.frm.submit()
	}
	
//function to show results according to string entered by user in the textbox
function fnShowResult()
{

var textSpeaker 
var ErrMsg,s
	ErrMsg = ""
	tcount = document.frm.tcount.value;
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
			
			document.frm.action="selectstore.asp?textval1="+textSpeaker+"&s=s"+"&p1="+ eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
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
	    tcount = document.frm.tcount.value;
	    theForm.SortBy.value = fldName;
	    //theForm.SortMsg.value = headerText;
	    theForm.SortColumn.value = columnNum;
	    theForm.action = "selectstore.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
	    theForm.submit( );
	}

function fnChecked(val,obj)
{
			
			var tcount 
			
			
			tcount = document.frm.tcount.value;
				
			if (obj.checked == true)
			{
					tcount = val;
					
			}

			window.document.frm.tcount.value = tcount;
}	

//function is called when confirms the deletion of the record
function fnOk(action)
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
								str = "document.frm.txtamt"+tcount+".value"
								value = eval(str)
								if (value==0)
								{
									alert("The selected store has left with no amount")
								}
								
								else
								{
									document.frm.SortBy.value=""
									document.frm.SortColumn.value=""
									document.frm.PriorSortBy.value=""
									document.frm.hidsortfield.value=""
									document.frm.hidsortcol.value=""
									document.frm.hidsortorder.value=""
									document.frm.hidpagechanged.value=""
									window.opener.document.forms[0].txtstoreid.value= tcount;
									window.opener.document.forms[0].txtamount.value= value;
								
									window.close()
								}	
			
				
		}
}
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">

<%
dim strid,strcustomers,showresult
showresult="N"

'Retriving the id of the reseller
if trim(Request.QueryString("action"))<>"" then
	strid=trim(Request.QueryString("action"))
end if
if trim(Request.Form("hidid"))<>"" then 
	strid = trim(Request.Form("hidid"))
end if
'Retriving the month & year
compdate1=trim(Request.Form("fromdate"))
compdate2=trim(Request.Form("enddate"))

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
		'Response.Write "NOT PAGE ---"&ordering

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
sqlid	= "select distinct(store_id) from sys_billing where ( store_id )not in (select fld_customer_id from tbl_request) group by  store_id order by "&ordering&"" 
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
					
		sqlid = "select distinct(store_id) from sys_billing where ( store_id )not in (select fld_customer_id from tbl_request) "&_
				" and store_id LIKE ('%"&scriteria1&"%') group by store_id  "&_
				"  order by "&ordering&""
		end if 
		
end if	

' if the user click on the SHOW all button 
if Request.QueryString("s") = "a" then
	 Searchval=""
	 showresult=""  
	 	sqlid	= "select distinct(store_id) from sys_billing where ( store_id )not in (select fld_customer_id from tbl_request)group by  store_id order by "&ordering&"" 
end if
'executing sql query
set rsID=server.CreateObject("ADODB.RecordSet")							
rsID.CursorLocation = 3
rsID.open sqlid,conn,2,2


%>
          <form name="frm" method="post" action="">
<table >
            <TABLE border=1 cellPadding=0 cellSpacing=0 width="100%" align=center>
              <TBODY>        
                  </TBODY></TABLE></TD>
          <TD height="15" vAlign="top" width="570"></TD></TR>
        <TR>
        &nbsp;
          <TD class=pagetitle height="400" vAlign="top" >
          <table border="0" id=TABLE1  align=center width="100%" >
  <TBODY>
          <tr>
                    <td width="15%" >

          Store Id</td>
          <td>
          <input name="txtname" >
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
       <td><b>Stores List</b></td>
       <td></td>
       <td></td>
       <td><b>No of Stores:<%=rsid.RecordCount%></b></td>
		<hr noshade>	
          </tr>
          
          <tr bgColor=#dddddd>
               
         <td width="5%"></td>
         <td width="30%">
					<strong><%=prefixes(1)%><A href="javascript:goSort('store_id',1);">Store ID</strong></A>
         </td>
			<td>
				Amount
			</td>
			<td>Reference</td>
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
						intcustomerid  = trim(rsID("store_iD"))
						sqlReseller = "SELECT fld_reseller_id from tbl_reseller_customer_master where fld_customer_id="&intcustomerid 
						set rsReseller = conn.execute(sqlReseller)
						if not  rsReseller.eof then 
							resellerID = trim(rsReseller("fld_Reseller_id"))
							resellerID = "By Reseller Id "&resellerID
						else
							resellerID= "Direct Referrance"
						end if
							
			 %>
          </tr>
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
          
          <tr>
          <td>
         <%
         if trim(tcount) = "" then 
				Keyname=trim(rsID("store_iD"))	
					Response.Write "&nbsp;<input type=radio "&_
					" name='Id' value='"&Keyname&"' "&_
					" onclick=javascript:fnChecked('"&Keyname&"',this)>"
			else		
					dim arrayofcount
					arrayofcount = split(tcount,",")

					for i = 0 to ubound(arrayofcount)-1
						if (cint(arrayofcount(i)))=	cint((Keyname)) then
								status = "Checked"
					exit for
						else
								status = ""		
						end if
								next 
											
						Response.Write "&nbsp;<input type=radio "&_
						" name='Id' "&status&" value='"&Keyname&"' "&_
						" onclick=javascript:fnChecked('"&Keyname&"',this)>"
						
			end if	%>
         </td>
          <td width="30%">
				<%= intcustomerid%>
			</td>
		
          <td width="30%">
				<%
				sql_select = "select sys_billing.amount,sys_payments.transdate,sys_payments.payment_term from sys_payments inner join sys_billing on sys_payments.store_id = sys_billing.store_id where sys_payments.store_id='"&intcustomerid&"' order by sys_payments.transdate desc"
				set rs_store = conn.execute(sql_select)
			
		
			if not rs_store.eof then
				Amount = rs_store("Amount")
			else
				Amount = 0
			end if
				txtamount = "txtamt"&intcustomerid
				Response.Write "$ "&amount
		%>
			<input type="hidden" value="<%=amount%>" name="<%=txtamount%>" >
		<%

				%>			
			</td>
			<td>
				<%=resellerID%>
			</td>
          </tr>
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
			 <td>
			 </td>
			<td><font color="red" face="arial" size="2"> <br>No Record Found</font>
			</td>
			<td>
			</td>
          </tr> 
		<%	
		end if
		
			%>
			
			<% rsGetData.close
			rsGetData=nothing%>
		<tr>
		  		<td>
					<input type="hidden" value="<% =strid%>" name="hidid">
						<INPUT Type="hidden" Name="hidSearch" Value="<%=Searchval%>">
						 <!--<br>SortBy-->     <INPUT Type="hidden" Name="SortBy">
<!--<br>SortColumn-->  <INPUT Type="hidden" Name="SortColumn">
<!--<br>PriorSortBy--> <INPUT Type="hidden" Name="PriorSortBy" Value="<%=priorOrder%>">
<!--<br>hidsortfield--><INPUT Type="hidden" Name="hidsortfield" value="<%=sortfield%>"> <!--second page and n th page navigation-->
<!--<br>hidsortcol-->  <INPUT Type="hidden" Name="hidsortcol" value="<%=sortcol%>"> <!--second page and n th page navigation-->
<!--<br>hidsortorder--><INPUT Type="hidden" Name="hidsortorder" Value="<%=sortorder%>"> <!--second page and n th page navigation-->
						<INPUT Type="hidden" Name="HidSearch" Value="<%=Searchval%>">
						<input type="hidden" Name="HidScriteria1" value="<%=scriteria1%>">
<!--<br>hidpagechanged--><INPUT Type="hidden" Name="hidpagechanged" Value="">

		<INPUT Type="hidden" Name="tcount" Value="<%=tcount%>">
					
				</td>
		
          	<tr>
           <tr>
			<td >&nbsp;
			</td>
          </tr> 
        <TBODY>
        <tr>
		<td>
		<INPUT name="OK" type="button" value="OK" onclick="javascript:fnOk()">
		<INPUT name="Close" type="button" value="Close" onclick="javascript:window.close()">
					 
		</td>	
				
		 </tr>
		  
         <tr>
				<td></td>	
				<td >  </td>
				<td> </td>
				<td align="right" valign='Right'><font face=arial size=2></font></td>			    
		  </tr></TBODY>
		  		
           
          </table>
          <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
             <tr>
		

			<td align=middle><font size="1" face="arial"  ><%DisplayNavBar p1,pageno,rsID.PageCount,10 %></font></td>
		
          </tr> 
          <% rsID.Close
          rsID=nothing%>
               <TR>
         

                <TD height=20></TD></TR></TABLE>
                
                
                
                
                
       </table>         
                
                
                
                
                
                
                </form>
              
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
