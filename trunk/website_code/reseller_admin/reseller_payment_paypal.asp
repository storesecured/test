<!--#include file = "include/ESCheader.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the email id of the reseller whose payment mode is paypal.
'	Page Name:		   	reseller_payment_paypal.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_list.asp
'	Output Page:	    reseller_payment_paypal.asp
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
dim sqlemail,intresellerid
if trim(Request.QueryString("action"))<> "" then 
	intresellerid= trim(Request.QueryString("action"))

	'Retriving the email of reseller
	sqlemail = "select fld_email_address_for_paypal from tbl_esc_reseller_payment_mode where fld_reseller_id="&intresellerid&""
	set rsemail = conn.execute(sqlemail)
	
	if not rsemail.eof then 
		Email = rsemail("fld_email_address_for_paypal")
	end if	
	rsemail.close
	rsemail=nothing
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
    <TD>    <!--#include file="incESCmenu.asp"-->
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
          <TD class=pagetitle height=400 vAlign=top>Payment Mode - Paypal
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action="" method=post name="frm">
                       <tr>
               <br>
				 <td>
					 <b> Paypal Email address</b>
				</td>
				<td>
					 
				</td>	
				</tr>	
              <tr>
               <br>
				 <td>
					  <br>
					  Email Address
				</td>
				<td>
					<%= Email%>
				</td>	
				</tr>	
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
