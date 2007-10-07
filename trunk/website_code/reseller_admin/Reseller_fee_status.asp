<!--#include file = "include/ESCheader.asp"-->
<%

Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the list of all the registered resellers.
'	Page Name:		   	reseller_fee_status.asp
'	Version Information:EasystoreCreator
'	Input Page:		    escadmin.asp
'	Output Page:	    reseller_fee_status.asp
'	Date & Time:		10 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel11=3
mSel12=2
%>
<HTML>
	<HEAD>
		<TITLE>Easystorecreator</TITLE>
		<META content=no-cache name=Pragma>
		<META content=no-cache http-equiv=pragma>
		<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
		<LINK href="images/style.css" rel=stylesheet type=text/css>
		<META content="MSHTML 5.00.3813.800" name=GENERATOR>
		<script language=javascript src="include/commonfunctions.js"></script>
		<script language="javascript">
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
					theForm.action = "reseller_fee_status.asp?p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
					theForm.submit( );
			}
			
			function fnShowAll()
			{
				tcount = document.frm.tcount.value;
				document.frm.hidpagechanged.value ="True"
						p1 = "0"
						p2 = "1"		
						p3 = "1"
					hitlistid = "1"  
			document.frm.action = "reseller_fee_status.asp?s=a"+"&p1=" + eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
			document.frm.submit()
			}
			//function to show results according to string entered by user in the textbox
			function fnShowResult()
			{

			var textSpeaker 
			var ErrMsg,s
				ErrMsg = ""
				document.frm.hidpagechanged.value = "True"
				tcount = document.frm.tcount.value;

						p1 = "0"
						p2 = "1"
						p3 = "1"
					hitlistid = "1"    

			// check whitespace
								
				if(isWhitespace(document.frm.txtsearch.value) == true)
					{	
							ErrMsg = ErrMsg + "Enter some criteria.\n"
					}
				else 	
					{
						textSpeaker = document.frm.txtsearch.value
					
						
					}	

			//check for numbers
			if (isWhitespace(document.frm.txtsearch.value) == false)
				{
					if( isAllNumeric(document.frm.txtsearch.value) == true)
						{	
							ErrMsg = ErrMsg + "Numbers are not allowed in the search criteria.\n";
						}
				}
			//check for special characters
				if (isWhitespace(document.frm.txtsearch.value) == false)
				{
				
					if( isAlphaNumeric(document.frm.txtsearch.value) == false)
						{	
							ErrMsg = ErrMsg + "Special characters are not allowed in the search criteria.\n"
						}
				
				}
				
				

			if (ErrMsg != "")
					{
						alert(ErrMsg)
					}
			else
					{
						
						document.frm.action = "reseller_fee_status.asp?textval1="+textSpeaker+"&s=s"+"&p1="+ eval(p1) + "&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
						document.frm.submit()
					}
			}

			function jsHitListAction(p1,p2,p3,hitlistid)
				{
				var str ="";
				tcount = document.frm.tcount.value;
					document.frm.hidpagechanged.value ="True"
					document.frm.SortBy.value = document.frm.hidsortfield.value
				document.frm.action = "reseller_fee_status.asp?p1="+ eval(p1)+"&p2=" +eval(p2)+ "&p3="+eval(p3)+"&hitlistid="+eval(hitlistid)+"&tcount="+tcount
				document.frm.submit()
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
								tempval = ""
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
			function fnsave()
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
									var ans = confirm("Are you sure to refund the license fee?")
									if (ans)
									{
											document.frm.SortBy.value = ""
											document.frm.SortColumn.value = ""
											document.frm.PriorSortBy.value = ""
											document.frm.hidsortfield.value = ""
											document.frm.hidsortcol.value = ""
											document.frm.hidsortorder.value = ""
											document.frm.hidpagechanged.value = ""
											document.frm.action = "reseller_fee_status.asp?action=action&tcount="+tcount
											document.frm.submit()	
							
									}
						}
			}

		</SCRIPT>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">

<%
 '*********************************CODE ADDED TO CHANGE THE STATUS HERE************
 if trim(Request.QueryString("action")) = "action" then
	tcount = trim(Request.QueryString("tcount"))
	tcount = mid(tcount,1,len(tcount)-1)
	sqlUpdate = "Update tbl_reseller_master set refund=1 where fld_Reseller_id in ("&tcount&")"
	conn.execute(sqlUpdate)
	sqlUpdate1 = "Update tbl_reseller_master set refund=0 where fld_Reseller_id not in ("&tcount&")"
	conn.execute(sqlUpdate1)
	StatusFlag = "1"
 end if 
'*********************************CODE ADDED TO CHANGE THE STATUS HERE************ 
dim showresult
 showresult="N"
 
 '''''''code for sorting'''''''''''''''


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
''''''''''''''''''''''''''''''''''''''

ordering = "" & Request("SortBy")
'*******************************************************

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

if Request.Form("hidpagechanged") <> "" then

	'these values are required when no sorting and only navigation takes place.	
		ordering = Request("hidsortfield")
		sortorder = Request.Form("hidsortorder")
		sortcol = Request.Form("hidsortcol")
		ordering = ordering & " " & sortorder
		
		if sortorder = "ASC" then
			prefixes(sortcol) = "<img src=asc.gif border=0 alt='sort order ascending'>"
		else
			prefixes(sortcol) = "<img src=desc.gif border=0 alt='sort order descending'>"
		end if		
else
	If ordering = priorOrder Then
		sortorder = "DESC"
	  ordering = ordering & " DESC"
	  priorOrder = ""
	  prefixes(sortColumn) = "<img src=desc.gif border=0 alt='sort order descending'>"
	Else
		sortorder = "ASC"
	  priorOrder = ordering 
	  prefixes(sortColumn) = "<img src=asc.gif border=0 alt='sort order ascending'>"
	End If
end if
dim sqlGetData,strfirst,strwebsite,strmail,sqlmode,intmode,intresellerid,rsgetdata
dim Searchval,Hidsearch
if Request.QueryString("tcount") <> "" then
	tcount = Request.QueryString("tcount")
end if

'Retriving the reseller information from the database
'sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,fld_reseller_id,refund,fld_license_fee from 'TBL_Reseller_Master order by "&ordering&""
'response.write"SQLgetdata="&SQLgetdata

sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,a.fld_reseller_id,refund,fld_license_fee , count(store_id) as count"
sqlGetData = sqlGetData & " from TBL_Reseller_Master a, store_settings b "
sqlGetData = sqlGetData & " where b.service_type>0 and b.trial_version=0 and b.reseller_id = a.fld_reseller_id"
sqlGetData = sqlGetData & " group by fld_first_name,fld_Last_Name,a.fld_reseller_id,refund,fld_license_fee "
sqlGetData = sqlGetData & " order by "&ordering&""


'sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,a.fld_reseller_id,refund,fld_license_fee , count(store_id) as count"
'sqlGetData = sqlGetData & " from TBL_Reseller_Master a, store_settings b "
'sqlGetData = sqlGetData & " where b.service_type>0 and b.trial_version=0 and b.reseller_id = a.fld_reseller_id"
'sqlGetData = sqlGetData & " group by fld_first_name,fld_Last_Name,a.fld_reseller_id,refund,fld_license_fee "
'sqlGetData = sqlGetData & " order by "&ordering&""
'response.write"sqlGetData="&sqlGetData

'response.end

if Request.QueryString("s") = "s" or Request.Form("HidSearch") = "search" then
	Searchval = "search"
	showresult = "y"
	
	'search the entire word in the list 	
		scriteria1 = trim(Request.QueryString("textval1"))
		scriteria1 = fixquotes(scriteria1)
		
		if scriteria1 = "" then 
			scriteria1 = trim(Request.Form("HidScriteria1"))
		end if 
		if scriteria1 <> "" then 	
				'sqlgetdata = "select fld_first_name+' '+fld_Last_Name as name ',fld_reseller_id,refund,fld_license_fee from TBL_Reseller_Master" &_
'			" WHERE fld_first_name LIKE ('%"&scriteria1&"%') order by "&ordering&" "
	sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,a.fld_reseller_id,refund,fld_license_fee , count(store_id) as count"
			sqlGetData = sqlGetData & " from TBL_Reseller_Master a, store_settings b "
			sqlGetData = sqlGetData & " where b.service_type>0 and b.trial_version=0 and b.reseller_id = a.fld_reseller_id"
			sqlGetData = sqlGetData & " and fld_first_name LIKE ('%"&scriteria1&"%') "
			sqlGetData = sqlGetData & " group by fld_first_name,fld_Last_Name,a.fld_reseller_id,refund,fld_license_fee "
			sqlGetData = sqlGetData & " order by "&ordering&""

			'sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,a.fld_reseller_id,refund,fld_license_fee , count(store_id) as count"
			'sqlGetData = sqlGetData & " from TBL_Reseller_Master a, store_settings b "
			'sqlGetData = sqlGetData & " where b.service_type>0 and b.trial_version=0 and b.reseller_id = a.fld_reseller_id"
			'sqlGetData = sqlGetData & " and fld_first_name LIKE ('%s%') "
			'sqlGetData = sqlGetData & " group by fld_first_name,fld_Last_Name,a.fld_reseller_id,refund,fld_license_fee "
			'sqlGetData = sqlGetData & " order by "&ordering&""
		end if 
end if	

' if the user click on the SHOW all button 
if Request.QueryString("s") = "a" then
	 Searchval = ""
	 showresult = ""  
	'	SQLgetdata = "select fld_first_name+' '+fld_Last_Name as name ,fld_reseller_id,refund,fld_license_fee from TBL_Reseller_Master order by "&ordering&" "		
	'sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,a.fld_reseller_id,refund,fld_license_fee , 'count(store_id) as count"
'	sqlGetData = sqlGetData & " from TBL_Reseller_Master a, store_settings b "
''	sqlGetData = sqlGetData & " where b.reseller_id = a.fld_reseller_id"
'	sqlGetData = sqlGetData & " group by fld_first_name,fld_Last_Name,a.fld_reseller_id,refund,fld_license_fee "
'	sqlGetData = sqlGetData & " order by "&ordering&""
	'response.write"<br>sshowall="&sqlGetData
sqlGetData = "select fld_first_name+' '+fld_Last_Name as name ,a.fld_reseller_id,refund,fld_license_fee , count(store_id) as count"
sqlGetData = sqlGetData & " from TBL_Reseller_Master a, store_settings b "
sqlGetData = sqlGetData & " where b.service_type>0 and b.trial_version=0 and b.reseller_id = a.fld_reseller_id"
sqlGetData = sqlGetData & " group by fld_first_name,fld_Last_Name,a.fld_reseller_id,refund,fld_license_fee "
sqlGetData = sqlGetData & " order by "&ordering&""
'	response.end	
end if

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
				<form name="frm" method="post" action="" >
		      <TABLE border=0 cellPadding=0 cellSpacing=0>
				    <TBODY>
							<TR>
								<TD class=title><B>Easystorecreator</B></TD>
								<TD class=special width=200>&nbsp;</TD>
								<TD align=left class=special width="60%"><UL><BR><BR></UL>
								</TD>
							</TR>
						</TBODY>
					</TABLE>
			</TD>
		</TR>            
		<TR>
			<TD><!--#include file="incESCmenu.asp"--></TD></TR>
		<TR>
			<TD>
	      <TABLE border=0 cellPadding=0 cellSpacing=0>
		      <TBODY></TBODY>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD>
		     <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
			    <TBODY>
					  <TR vAlign=top>
							<TD rowSpan=2 width=180>
							  <TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
									<TBODY></TBODY>
									</TABLE>
							</TD>
							<TD height=15 vAlign=top width=570></TD>
						</TR>
						<TR>
							<TD height=400 vAlign=top>
			         <table border="1" width="100%" > 
						    <tr>
									<td>Search</td>
			          <tr>
									<td>First Name</td>
				          <td><input type="textbox" name="txtsearch" maxlength=12></td>
			          </tr>
								<tr>
									<td></td>
									<td><input type="button" name="search" value="Show" onclick="javascript:fnShowResult()">
									<%if showresult = "y" then %>
									<input type="button" name="showall" value="Show all"  onclick="javascript:fnShowAll()">
									<%end if%>
									</td>
				          Resellers List 
									<hr noshade>	
						   </tr>
								<tr></tr>
								<tr>
														<!-- <td width ="20%" height="20" class="bodyboldlink"><strong><%=prefixes(2)%><a href="javascript:goSort('fldAuthor',2);"><%=getstring("lblAuthor")%></a></strong></td>-->
										<td width="23%"><div align="center"></div></td>
								</tr>         
			          <tr bgColor=#dddddd>			
								 <td ><strong><%=prefixes(1)%><a href="javascript:goSort('fld_first_name',1);">Reseller Name</strong></a>
								 </td>
								<td >
									<%=prefixes(2)%><a href="javascript:goSort('count',2);">Customers Referred</a>
								</td>
								<td>
									<%=prefixes(3)%><a href="javascript:goSort('fld_license_fee',3);">License Fee Paid</a>
								</td>			
								<td>
									Status
								</td>			
		          </tr>
				      <tr>
								<td >&nbsp;</td>
							</tr> 
	<%

	
				if not rsGetData.eof then
				' FOR paging
											' -------------------------------		
												dim pageno,p1,curpage
												ctrNum = 1
												pageno = trim(request.querystring("p2"))
																								
												rsGetData.pagesize = 50
												p1=Request.QueryString("p1") 
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
												iNum = pageNo - 1
												ctrNum = (iNum*rsGetData.PageSize) + 1
												
												
												curpage = 0
	                 							' ---------
						 do while curpage < rsGetData.PageSize
							strfirst = Trim(rsGetData("name"))
							RefundStatus = Trim(rsGetData("refund"))
							Fee = Trim(rsGetData("fld_license_fee"))
							Fee = FormatNumber(Fee,2)
							NoCust = 0
							NoCust = Trim(rsGetData("count"))
							
							intresellerid = trim(rsGetData("fld_reseller_id"))
							sqlCust =  " select IsNull(count(store_id),0) as count from store_settings where service_type>0 and trial_version=0 and reseller_id="&intresellerid&""
										
							'response.write"sqlCust="&sqlCust
							'response.end
							set rsCust=conn.execute(sqlCust)
							if not rsCust.eof then
'								NoCust = trim(rsCust("count"))
								
							end if
			%>
          <tr>
			<td >
			<%=strfirst%> 		
			</td>
			<td >
			<%=NoCust%>
			<td >$
			<%=Fee%>
			</td>
			
			
			</td>
			<td >
			
			 <%
			if RefundStatus = 1 then 
				status = "Checked"
				if instr(1,tcount,rsGetData("fld_reseller_id")) = 0 then 
					tcount = rsGetData("fld_reseller_id")&","&tcount
				end if
				
				
			end if
			
			if trim(tcount) = "" then 
			'Response.Write "Status"&rsGetData("fld_reseller_id")&"-"&Status
			
					Response.Write "&nbsp;<input type=checkbox class=INPUT_check border=0 "&_
					" name='Id'  value='"&rsGetData("fld_reseller_id")&"' "&_
					" onclick=javascript:fnChecked('"&rsGetData("fld_reseller_id")&"',this)>"
			else		
					dim arrayofcount
					arrayofcount = split(tcount,",")

					for i = 0 to ubound(arrayofcount)-1
						if (cint(arrayofcount(i)))=	cint((rsGetData("fld_reseller_id"))) then
								status = "Checked"
					exit for
						else
								status = ""		
						end if
								next 
											
						Response.Write "&nbsp;<input type=checkbox class=INPUT_check border=0 "&_
						" name='Id' "&status&" value='"&rsGetData("fld_reseller_id")&"' "&_
						" onclick=javascript:fnChecked('"&rsGetData("fld_reseller_id")&"',this)>"
			end if			
%>
			</td>
			
          </tr>
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
			
			
			else
			if Request.QueryString("s") <> "" or Request.Form("HidSearch") = "search" then
			%>
			<tr>
			<td></td>
			<td><font color="red">No records found.</font></td>
			<td><font color="white"></font></td>
			</tr>
			<%end if
			end if
			 %>

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

            <tr>
            <td>
            <td>
			<td >
			
			</td>
			<td><td>
          </tr> 
        <tr>
        <td >
        <input type="button" value="Save" name="Save" onclick="javascript:fnsave()">
        <input type="button" value="Back" name="back" onclick="javascript:history.back()">
        </td>
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
        
			<td align="middle"><font size="1" face="arial" ><%DisplayNavBar p1,pageno,rsGetData.PageCount,10 %></font></td>
		<% rsgetdata.Close
		rsGetData=nothing%>
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
End sub

if StatusFlag = "1" then
%>
<script language="javascript">
	alert("The License fee refund status has been updated")
	document.location.href = "reseller_fee_status.asp"
	
	
</script>
<%end if%>
