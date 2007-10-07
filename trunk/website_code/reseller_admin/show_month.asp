<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page checks login ID and password for the user for the resellers site
'	Page Name:		   	reseller_login.htm
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_login.htm
'	Date & Time:		13 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel3=3
mSel4=2
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="../include/commonfunctions.js"></SCRIPT>
<script language=javascript>
function fnshow()
{
	var ErrMsg
	ErrMsg="";
	
	if(document.frmshowmonths.frommm.value==0)
		{
		ErrMsg=ErrMsg + "From month is mandatory.\n";
		}
			
	if(document.frmshowmonths.fromyy.value==0)
		{
		ErrMsg=ErrMsg + "From year is mandatory.";
		}
	if(ErrMsg !="")
		{
		alert(ErrMsg);
		}
	else
		{
		document.frmshowmonths.action="update_payment_status.asp?action=check";
		document.frmshowmonths.submit();
		}
}

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%

'code to retrieve resellerid
dim strid
strid=trim(Request.QueryString("action"))

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
       <!--#include file="incESCmenu.asp"   --></TD></TR>
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
         <form name="frmshowmonths" method=post >
         <input type="hidden" value="<%=strid%>" name="hidid">
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>     
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1">
          <tr>
          Reseller Update Payment
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
       <%
       'code here to rerieve the years fro the table and display 
		Set rsYear = Server.CreateObject("ADODB.Recordset")
		With rsYear
			.Source = "Get_Year"
			.ActiveConnection = strConn
			.CursorType = adOpenForwardOnly
			.LockType = adLockReadOnly
			.Open
		End With
		
		
		
		if not rsYear.EOF then 
			while not rsYear.EOF 
			YearId = rsYear("fld_Year_Id")
			Yearname = rsYear("fld_Year_Name")
		
		
		 
       %>
       
             <option value="<%= YearId%>"><%= Yearname %></option>  
        <%
				rsYear.MoveNext		
				wend
				
		end if	
			Set rsYear = Nothing
		%>
          
    </select>
       
   
          <tr>
			<td >&nbsp;
			</td>
          </tr> 
          <tr>
				
				<td>
					 &nbsp;&nbsp;<INPUT name="save" type="button" value="Update" onclick="javascript:fnshow()"> 
					 <input type="reset" value="Reset" name=reset >
				</td>	
				
		  </tr>
          
          </table></FORM>  
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%" style="LEFT: 0px; TOP: 178px">
              <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE>
</BODY></HTML>
