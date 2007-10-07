<!--#include file = "include/ESCheader.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the information of the reseller whose payment mode is check.
'	Page Name:		   	reseller_payment_check.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_list.asp
'	Output Page:	    reseller_payment_check.asp
'	Date & Time:		10 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel3=3
mSel4=2
%>
<HTML><HEAD><TITLE>EasyStoreCreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language="javascript">

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
dim sqlGetData,intresellerid
if trim(Request.QueryString("action"))<>"" then
	intresellerid=trim(Request.QueryString("action"))

'Retriving information about reseller whose payment mode is check
sqlGetData = "select fld_payble_to,fld_address,fld_city,fld_state,fld_country,fld_zipcode from TBL_ESC_Reseller_payment_mode where fld_reseller_id="&intresellerid&""
set rsGetData = conn.execute(sqlGetData)
	if not rsGetData.eof then 

		strpayble = trim(rsGetData("fld_payble_to"))
		straddress = trim(rsGetData("fld_address"))
		strcity = trim(rsGetData("fld_city"))
		strstate = trim(rsGetData("fld_state"))
		strcountry = trim(rsGetData("fld_country"))
		intzipcode = trim(rsGetData("fld_zipcode"))
	end if
	'Retriving information about country
	sqlcountry="select country from sys_countries where country_id="&strcountry
	
	set rscountry=conn.execute(sqlcountry)
	if not rscountry.eof then
	strcountry1=trim(rscountry("country"))
	end if
	rscountry.close
	rscountry=nothing
	
end if
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
          <TD class=title><B>EasyStoreCreator</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%">
            <UL><BR><BR></UL></TD></TR></TBODY></TABLE></TD></TR>
              
  <TR>
    <TD><!--#include file="incESCmenu.asp"-->
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
          <TD class=pagetitle height=400 vAlign=top>Payment Mode - Check
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action="" method=post name="frm">
              <tr>
               <br>
				 <td>
					 <b>Address for sending check </b>
				</td>
				<td>
					 
				</td>	
				</tr>	
				
            <tr>
             <br>
				<%
				if not rsGetData.eof then
				%>
				 <td>
					  Payable to 
				</td>
				<td>
					<%=strpayble%>
				</td>	
			</tr>	
			<tr>
            	 <td>
					  Flat no 90
				</td>
				<td>
					 <%=straddress%>
				</td>	
			</tr>	
			<tr>
            	 <td>
					  State
				</td>
				<td>
					 <%=strstate%>
				</td>	
			</tr>	
			<tr>
            	 <td>
					  City
				</td>
				<td>
					 <%=strcity%>
				</td>	
			</tr>	
					
			<tr>
			
            	 <td>
					  ZipCode
				</td>
				<td>
					 <%=intzipcode%>
				</td>	
			</tr>	
			<tr>
			
            	 <td>
					  Country
				</td>
				<td>
					 <%=strcountry1%>
				</td>	
			</tr>	
				<%
				 rsGetData.movenext
			     end if
				%>
				<% rsGetData.close
				rsGetData=nothing
				%>
			<tr>
				<td>
				
				</td>
				
				<td>
				
				<br>
				           	
					 <INPUT name="back" type="button" value="Back" onclick="javascript:history.back()"> 
					 
				</td>	
				
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
