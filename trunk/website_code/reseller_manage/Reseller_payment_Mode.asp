<!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the payment mode for the reseller site.
'	Page Name:		    reseller_payment_Mode.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_payment_Mode.asp
'	Date & Time:		5 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel7=3
mSel8=2
%>

<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language="javascript">
function fnSelect()
{

//To check what user select from the list
var val
val =  document.frmmode.selkeyword.value
	if( val==0)
	{
			alert("Please select mode of Payment.")
	}
	if( val==1)
	{
		document.location.href = "Reseller_payment_Emailaddress.asp"
	}
	if( val==2)
	{
		document.location.href = "Reseller_payment_address.asp"
	}

}
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
dim strGetemail ,intResellerId

'Retriving the reseller id from session variable
intResellerId=session("ResellerId")

'Retriving the payment mode for the reseller
strmode="select fld_payment_mode from tbl_esc_reseller_payment_mode where fld_reseller_id="&intResellerId&" "
set rsmode=conn.execute(strmode)
if not rsmode.eof then
	mode=trim(rsmode("fld_payment_mode"))
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
         <%
'***********************Code Added on 19th Jan 2005******************
%>
				<TR>
					<TD class=meniu height=20 onmouseout="style.backgroundColor='#EBF9D8', 			style.color='#8FB25E'" onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
					<A class=b href="minimum_balance.asp">Set Minimum Balance</A>
					</TD>
        </TR>   
  <%
	'****************************************************
	%>                      
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Choose Payment Mode
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action="" method=post name="frmmode">
              <tr>
               <br>
				 <td>
					  Select Payment Mode 
				</td>
				<td>
					 <select name="selkeyword" >
					<option value="0"  >--Select--</option>
					<option value="1" <% if mode="1" then %>selected<%end if%> >Paypal</option>
					<option value="2" <% if mode="2" then %>selected<%end if%> >Check Payments</option>  
					 </select>
				</td>	
				</tr>	
				<tr>
				<td>
					  
				</td>
				<td>
				<br>
				           
					 <INPUT name="Next" type="button" value="Next" onclick="javascript:fnSelect()"> 
					 <input type="reset" type="reset" value="Reset">
				</td>	
              </tr>
              
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
