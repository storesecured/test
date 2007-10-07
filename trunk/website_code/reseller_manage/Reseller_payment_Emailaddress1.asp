<!--#include file = "../include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the payment mode as paypal for the reseller site.
'	Page Name:		   	reseller_payment_Emailaddress.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_payment_Emailaddress.asp
'	Date & Time:		6 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
%>


<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="../include/commonfunctions.js" type=text/javascript></SCRIPT>
<script language="javascript">

function fnsave()

{
	var ErrMsg
	ErrMsg=""

	//Validations for blank email id
	if (isWhitespace(document.frmemailaddress.txtemail.value)== true)
	{
			ErrMsg = ErrMsg +"Email field should not be blank\n";
	}

	if (isWhitespace(document.frmemailaddress.txtemail.value)== false)
	{
		if(IsEmail(document.frmemailaddress.txtemail.value)==false)
		{
			ErrMsg = ErrMsg + "Invalid Email id \n" ;		
		}	
	}
	if (ErrMsg!="")
	{
	alert(ErrMsg);
	return false
	}
	else
	{
			document.frmemailaddress.hidsubmit.value="y";
			return true
	}	
}
</script>
</HEAD>

<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
dim email,strinsert,intResellerID
'retreving the reseller id from the session
'intResellerID = session("")
intResellerID =1

'Retriving the email id
if trim(request.querystring("hidsubmit"))="y"  then 
	email=Request.Form("txtemail")
	Response.Write "email1 ="&email
	'Inserting the email id into the table
	strinsert="insert into TBL_Esc_Reseller_payment_mode(fld_reseller_id,fld_payment_mode,fld_email_address_for_paypal)values("&intResellerID&",1,'"&email&"')"
	conn.execute (strinsert)
end if
%>
<DIV id=overDiv style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
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
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
          <TD class=menu noWrap><A class=menu
            href="reseller_home.asp">Customize Sales Website</A></TD>
      <TD class=menu noWrap><A class=menu 
            href="../default.asp" target=_blank
            >Preview</A></TD>
        <!--  <TD class=menu noWrap><A class=menu 
            href="Reseller_Admin.asp">Customize Administrative website</A></TD>-->
          <TD class=menu noWrap><A class=menu 
            href="Resellers_Customer_list.asp"
            >Customer List</A></TD>
             <TD class=menu3 noWrap><A class=menu2 
            href="Resellers_Sales.asp">My Sales History</A></TD>
          <!--<TD class=menu noWrap><A class=menu 
            href="Resellers_Customer_list.asp">My 
            Account</A></TD>-->
           <!--   <TD class=menu noWrap><A class=menu 
            href="Resellers_Amount.asp">Amount Due</A></TD>-->
          <!--<TD class=menu noWrap><A class=menu 
            href="http://manage.abetterwaystore.com/support_list.asp">Support 
            Request</A></TD>-->
          <TD class=menu noWrap><A class=menu 
            href="logout.asp">Logout</A></TD>
          <TD align=right class=menu noWrap width="100%"><FONT 
            color=white>Reseller Id #: <B>1</B></FONT></TD></TR></TBODY></TABLE></TD></TR>
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
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="Reseller_sales_history.asp">View Sales History
              </A></TD>
        </TR>
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="Reseller_payment_Mode.asp">Choose Payment Mode
              </A></TD>
        </TR>   
            
                          
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Payment Mode - Paypal
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 width="90%">
              <FORM method=post name="frmemailaddress" action="javascript:fnsave()">
              
              <input type="text" value="" name="hidsubmit">
                       <tr>
               <br>
				 <td>
					 <b> Enter the Paypal Email address</b>
				</td>
				<td>
					 
				</td>	
				</tr>	
              <tr>
               <br>
				 <td>
					  Email Address
				</td>
				<td>
					 <input type="text" name="txtemail">
				</td>	
				</tr>	
				<tr>
				<td>
					  
				</td>
				<td>
				<br>
				           	
					 <INPUT name="Back" type="button" value="Back" onclick="javascript:document.location.href='reseller_payment_mode.asp'"> 
					 <INPUT name="Save" type="Submit" value="Save" > 
					 <input type="reset" type="reset" value="Reset">
				</td>	
              
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
