<!--#include file = "include/header.asp"-->
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
mSel7=3
mSel8=2
%>


<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
<LINK href="images/style.css" rel=stylesheet type=text/css>
<script language="javascript" src="../include/commonfunctions.js"></script>

<script language="javascript">
function fnsave()

{
	var ErrMsg
	ErrMsg=""

	//Validations for blank  & invalid email id
	if (isWhitespace(document.frmemailaddress.txtemail.value)== true)
	{
			ErrMsg = ErrMsg +"Email field is mandatory.\n";
	}

	if (isWhitespace(document.frmemailaddress.txtemail.value)== false)
	{
		if(IsEmail(document.frmemailaddress.txtemail.value)==false)
		{
			ErrMsg = ErrMsg + "Invalid Email id. \n" ;		
		}	
	}
	if (ErrMsg!="")
	{
	alert(ErrMsg);
	
	}
	else
	{
		document.frmemailaddress.action ="Reseller_payment_emailaddress.asp?action=save";
		document.frmemailaddress.submit();
	}	
}
</script>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
dim email,strinsert,intResellerID,stremail,intFlag
intFlag =0
'retreving the reseller id from the session
intResellerID = trim(session("ResellerID"))

'Retriving the email id
'check here if the reseller has already made any entry for paypal(hardcode check for payment mode=1"
'-------------------------------------------------------------------------------
stremail = " select fld_email_address_for_paypal from tbl_esc_reseller_payment_mode where fld_reseller_id="&trim(session("ResellerID"))
		   
set rsemail=conn.execute(stremail)
reseemail=trim(rsemail("fld_email_address_for_paypal"))
reseemail.close
set reseemail=nothing

if trim(request.querystring("action"))="save"  then 
	'check here if the reseller has already made any entry paypal or for check(irrespective of payment mode)
	'-------------------------------------------------------------------------------
	email=trim(Request.Form("txtemail"))
	strcheck = " select fld_email_address_for_paypal from tbl_esc_reseller_payment_mode where fld_reseller_id="&intResellerID&""
	set rscheck = conn.execute (strcheck)  
	if not rscheck.eof then
		emailcheck=trim(rscheck(checkencode("fld_email_address_for_paypal")))
	
		'Entry is there

			'Update query for that reseller
			'Updating the email id into the table
		strupdate="Update TBL_Esc_Reseller_payment_mode set fld_email_address_for_paypal='"&email&"'"&_
							",fld_payment_mode=1 where fld_reseller_id="&intResellerID&""
		conn.execute (strupdate)
		intFlag = "1"
		else
		
	
	'Entry is not there
		' insert query for that reseller
	
	'Inserting the email id into the table
	strinsert="insert into TBL_Esc_Reseller_payment_mode(fld_reseller_id,fld_payment_mode,fld_email_address_for_paypal)"&_
			  " values("&intResellerID&",1,'"&email&"')"
	conn.execute (strinsert)
	intFlag = "2"
	end if
	rscheck.close
set rscheck=nothing
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
          <TD class=title><B>Reseller</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%">
            <UL><BR><BR></UL></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD>
      <!--#include file="incmenu.asp"-->
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
           
      <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="edit_reseller.asp">Edit Profile
              </A></TD>
        </TR>   
        
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top> 
Payment Mode - Paypal  

            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              
              <FORM action="" method=post name="frmemailaddress">
              <TBODY>
              <tr>
               <br>
				 <td>
					 <B>Enter the Paypal Email address</B> 
				</td>
<!--Entry is there Retreive the old and dispaly-->
				
				<td>
					 <INPUT type="textbox"name="txtemail" maxlength="50" value="<%= rsemail("fld_email_address_for_paypal") %>"> 
				</td>	
				
				</tr>	
				
              <tr>
                
                <td>
					  
				</td>
				<td>
					 
					        	
					 <INPUT name="Save" type="button" value="Save" onclick="javascript:fnsave()"> 
					 <input type="reset" type="reset" value="Reset">
					 <INPUT name="Back" type="button" value="Back" onclick="javascript:document.location.href='reseller_payment_mode.asp'"> 
				</td></tr>	
              
              <TBODY>
              <TR>
				

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE>
              
              </FORM>
</BODY></HTML>
<% if intFlag ="1" then%>
<script language="javascript">
	alert("Your email address has been updated.")
	document.location.href = "Reseller_payment_emailaddress.asp"
</script>
<%end if%>

<% if intFlag ="2" then%>
<script language="javascript">
	alert("Your email address has been added.")
	document.location.href = "Reseller_payment_emailaddress.asp"
</script>
<%end if%>
