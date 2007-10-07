<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page checks login ID and password for the user for the resellers site
'	Page Name:		   	update_payment_status.htm
'	Version Information: EasystoreCreator
'	Input Page:		    update_payment_status.asp
'	Output Page:	    update_payment_status.asp
'	Date & Time:		14 Aug 2004 		
'	Created By:			Rashmi Badve
'	Code added by Devki Anote for displaying the things properly
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
<SCRIPT language=JavaScript src="../include/commonfunctions.js" ></script>
<SCRIPT language=JavaScript src="images/script.js" type=text/javascript></SCRIPT>
<script language="javascript">
function fnsave()
{
var ErrMsg
ErrMsg="";

if (isWhitespace(document.frmupdate.txtamt.value)==true)
	{
	ErrMsg=ErrMsg + "Amount paying is mandatory.\n";
	}

if (isWhitespace(document.frmupdate.txtamt.value)==false)
{
	if (isAllNumeric(document.frmupdate.txtamt.value)==false)
	{
		ErrMsg=ErrMsg + "Only numbers are allowed.\n";
	}
	else if(eval(document.frmupdate.txtamt.value) > eval(document.frmupdate.hidbalance.value))
	{
		ErrMsg=ErrMsg + "Amount to pay should not be greater than amount balance.";
	}
	
}
if (ErrMsg!="")
	{
	alert(ErrMsg);
	}
else
	{
	document.frmupdate.action="update_payment_status.asp?action=save"
	document.frmupdate.submit();
	}
}

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
'asp code here
dim strmonth,sqlselect,rsselect,strid,stramttopaid,sqlgetbalance,rsgetdata,stramtpaid,stramtbalance
dim sqlpaymentmode,rsmode,stryear
'code here to reterive the values passed from the prev page
		strmonth=trim(Request.Form("frommm"))
		stryear=trim(Request.Form("fromyy"))

'retriving value of month  from hidden variable		
		if trim(Request.Form("hidmonth")) <> "" then 
				strmonth=trim(Request.Form("hidmonth")) &"<br>"
		end if	
		
'retriving value of year from hidden variable		
		if trim(Request.Form("hidyear")) <> "" then 
				stryear=trim(Request.Form("hidyear"))
		end if		

'retriving value of resellerid from hidden variable		
		if trim(Request.form("hidid"))<> "" then 
			strresellerid=trim(Request.form("hidid"))
		end if	
		
		if trim(Request.form("hidamt"))<> "" then 
			stramttopaid =  trim(Request.form("hidamt"))
		end if	


'query for retriving amount to pay to reseller from tbl_reseller_customer_master
if trim(Request.QueryString("action"))= "check" then 

	sqlselect = "select count(distinct(fld_customer_id)) as Refcount, sum(fld_amount_pay_to_reseller) as sum1 from "&_
			   " tbl_reseller_customer_master where fld_reseller_id="&strresellerid&" and "&_
			   " month(fld_plan_transaction_date)="&strmonth&" and "&_
			   " year(fld_plan_transaction_date)="&stryear&""
		set rsselect = conn.execute (sqlselect)
		'Response.Write "sqlselect"&sqlselect
		
		if not rsselect.eof then 
			count = trim(rsselect("Refcount"))
			stramttopaid=trim(rsselect("sum1"))
			stramtbalance =stramttopaid
		
		end if	

	if trim(count) = "0" then 
		%>
		<script language=javascript>
		alert("Currently there are no customers for the reseller");
		document.location.href = "reseller_list.asp"
		</script>
		<%	
	
	end if
	
	if isnull(sum1)= False then 
			stramttopaid=trim(rsselect("sum1"))
			stramtbalance =stramttopaid
	else

			intFlag = "3"

	end if	
	
				rsselect.close
				set rsselect=nothing

			
end if 

	
'query for retriving amount paid and balance from tbl_esc_reseller_payment

	sqlgetbalance="select fld_reseller_id,sum(fld_amount_paid) as paidamt,sum(fld_amount_balance) as balanceamt from tbl_esc_reseller_payment where fld_reseller_id="&strresellerid&" and fld_month='"&strmonth&"' and fld_year='"&stryear&"' group by fld_reseller_id " 
	set rsgetdata=conn.execute(sqlgetbalance) 
	'if isnull(paidamt)= False and isnull(balanceamt)= False then
	if not rsgetdata.eof then 	
	
		stramtpaid=trim(rsgetdata("paidamt"))
		stramtbalance=trim(rsgetdata("balanceamt"))
		'Response.Write "stramttopaid"&stramttopaid
		stramtbalance = CSng(stramttopaid)-CSng(stramtpaid)
	else
		
	end if
		rsgetdata.close
		set rsgetdata=nothing

'query for retriving payment mode from tbl_reseller_payment_mode
	sqlpaymentmode="select fld_payment_mode from tbl_esc_reseller_payment_mode where fld_reseller_id="&strresellerid&""
	set rsmode=conn.execute(sqlpaymentmode)
	
	if not rsmode.eof then
		strmode=trim(rsmode("fld_payment_mode"))
	end if

		rsmode.close
		set rsmode=nothing
	
'query for checking if there are recorde for particular reseller
dim sqlcheck,sqlcheckrecord,rscheck,rscheckrecord,amount,totalamount,sqlupdate,rsupdate,checkmonth,checkyear

		
if trim(Request.QueryString("action"))="save" then
	
    'retriving value of month  from hidden variable		
		checkmonth=cint(trim(Request.Form("hidmonth")))
'		 Response.Write "checkmonth="&checkmonth &"<br>"
		 
	'retriving value of resellerid from hidden variable		
	strresellerid=trim(Request.form("hidid"))
	
		
	'retriving value of year from hidden variable		
		checkyear=cint(trim(Request.Form("hidyear")))

		
	'getting value from (textbox) amount that the esc is paying to the reseller
		amount=trim(Request.Form("txtamt"))

		
	'retriving value of total amount to be paid from hidden variable(stramttopaid to reseller)	
		totalamount=trim(Request.Form("hidamt"))

		
	'retriving value of total amount paid from hidden variable(stramtpaid uptil now)	

		stramtpaid=trim(Request.Form("hidpaid"))

		
	
		stramtbalance=trim(Request.form("hidbalance"))
		
	'query for checking whether there is any record for the reseller or not 
		
		set RsObjcheck = server.CreateObject("adodb.recordset")
		strlclQC ="select fld_esc_reseller_payment_id,fld_amount_paid from TBL_ESC_RESELLER_PAYMENT WHERE fld_reseller_id =" &strresellerid&" and fld_month=" & checkmonth&"  and fld_year="& trim(checkyear)

		set rscheckrecord = conn.execute(strlclQC)
	
		
		
	if not rscheckrecord.eof then
		cuid=trim(rscheckrecord("fld_amount_paid"))
		stramtbalance  = CSng(stramtbalance)- CSng(amount)
		stramtpaid= CSng(stramtpaid) + CSng(amount)
	   
		'query for updating records
	   
	   strupdatequery= "update TBL_ESC_Reseller_Payment set fld_amount_paid="&stramtpaid&",fld_amount_balance="&stramtbalance& " " &_
   	   " where fld_reseller_id="&strresellerid&" and fld_month='"&checkmonth&"' and fld_Year='"&checkyear&"' "   	   
 
		conn.execute strupdatequery
		intFlag  = "2"
else
			'if stramtpaid="" or stramtpaid="0" then
			'retriving balance amount for the first time

				stramtbalance=cSng(totalamount)- cSng(amount)

				stramtpaid=0
				stramtpaid=cSng(stramtpaid)+ cSng(amount)


	
			sqlinsert=	"insert into tbl_esc_reseller_payment (fld_amount_paid,fld_amount_balance,fld_reseller_id,fld_month,fld_Year) values "&_
						" ("&stramtpaid&","&stramtbalance&","&strresellerid&","&checkmonth&","&checkyear&")"
			conn.execute(sqlinsert)

			intflag = "1" 

			'end if		

end if
	
		rscheckrecord.close
		set rscheckrecord=nothing
	
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
    <TD>
     <!--#include file="incESCmenu.asp"   --></TD></TR> <TR>
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
          <TD class=pagetitle height=400 vAlign=top>Update Payment Status For Month <%=monthname(strmonth)%>
            
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
            <FORM action="" method=post name="frmupdate">
                <% if intFlag ="3" then%>
        
              <tr>
                <td>
               <font color="red"> No Records Available for the Reseller</font>
				</td>
				<td>
				
				</td>	
				</tr>	
		    <% else %>
        
             <input type="hidden" name="hidamt" value="<%=stramttopaid%>">
			<input type="hidden" name="hidmonth" value="<%=strmonth%>">
			<input type="hidden" name="hidyear" value="<%=stryear%>">
			<input type="hidden" name="hidpaid" value="<%=stramtpaid%>">
			<input type="hidden" name="hidbalance" value="<%=stramtbalance%>">
            
            <tr>
               <br>
				 <td>
					 <b> Payment Status</b>
				</td>
				<td>
					 
				</td>	
				</tr>	
              <tr>
               <br>
				 <td>
					  <br>
					 Net Profit
				</td>
				<td>
					 $ <%=stramttopaid%>
				</td>	
				</tr>	
				<tr>
               
				 <td>
					  <br>
					  Amount Balance
				</td>
				<td>
					 $ <%if trim(stramtbalance)="" or trim(stramtbalance)="0" then Response.Write "0" else response.write stramtbalance end if %>
					 
				</td>	
				</tr>	
				<tr>
               
				 <td>
					  <br>
					  Amount Paid
				</td>
				<td>
				<!--$ <%=amountpaid%>-->
					 $ <%if trim(stramtpaid)="" or trim(stramtpaid)="0" then Response.Write "0" else response.write stramtpaid end if%>
					 <%'=stramtpaid%>
					 
				</td>	
				</tr>
				<tr>
               
				 <td>
					  <br>
					  Pay 
				</td>
				<td>
					 $ <input type="text" name="txtamt">
					 
				</td>	
				</tr>	
				<tr>
               
				 <td>
					  <br>
					  Payment by
				</td>
				<td>
					 <% if strmode=2 then%>
					 By check 
					 <% else %>
					 By Paypal
					 <%end if%>
					 
				</td>	
				<%end if%>
				</tr>
					
				       <input type="hidden" value="<%=strresellerid%>" name="hidid">
				<tR><td>
					  
				</td>
				<td>
				<br>
				     <%if intflag<> "3" then%><input name="Update" type="button" value="Update" onclick="javascript:fnsave()"><%end if%> 
					 <INPUT name="back" type="button" value="Back" onclick="javascript:history.back()"> 
			
				</td>	
              </tr>
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>

<%if intflag = "2" then %>
	<script language="javascript">
		alert("Payment status has been updated.");
		document.location.href = "reseller_List.asp"
		
		
	</script>
<% end if%>

<%if intflag = "1" then %>
	<script language="javascript">
		alert("Payment status has been added.");
		document.location.href = "reseller_List.asp"
	</script>
<% end if%>