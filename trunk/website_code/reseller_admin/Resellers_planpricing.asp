<!--#include file = "include/ESCheader.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page displays the plan pricing for the reseller depending 
'						on his user range.
'	Page Name:		   	Resellers_planpricing.asp
'	Version Information:EasystoreCreator
'	Input Page:		    Resellers_planpricing.asp
'	Output Page:	    Resellers_planpricing.asp
'	Date & Time:		11 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	##############
msel5=3
msel6=2 
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="include/commonfunctions.js" ></script>
<script language=javascript>
function fnsave()
{
	var ErrMsg 
	ErrMsg=""
	
	//Validations for bronze plan prizing for blank fields.
	if (isWhitespace(document.frmplan.text11.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Bronze for range <0-50> is mandatory. \n" ;		
		}
		
	if (isWhitespace(document.frmplan.text11.value)== false)
		{
			if (isAllNumeric(document.frmplan.text11.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Bronze for range <0-50> should be numeric. \n" ;		
				}	
		}
		
	if (isWhitespace(document.frmplan.text12.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Bronze for range <50-100> is mandatory. \n" ;		
		}	
	
	if (isWhitespace(document.frmplan.text12.value)== false)
		{
			if (isAllNumeric(document.frmplan.text12.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Bronze for range <50-100> should be numeric. \n" ;		
				}	
		}
	if (isWhitespace(document.frmplan.text13.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Bronze for range <100 & above> is mandatory. \n" ;		
		}	

	if (isWhitespace(document.frmplan.text13.value)== false)
		{
			if (isAllNumeric(document.frmplan.text13.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Bronze for range <100 & above> should be numeric. \n" ;		
				}	
		}
	
	//Validations for Silver plan prizing for blank fields.
	if (isWhitespace(document.frmplan.text21.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Silver for range <0-50> is mandatory. \n" ;		
		}
		if (isWhitespace(document.frmplan.text21.value)== false)
		{
			if (isAllNumeric(document.frmplan.text21.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Silver for range <0-50> should be numeric. \n" ;		
				}	
		}
	
		
	if(isWhitespace(document.frmplan.text22.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Silver for range <50-100> is mandatory. \n" ;		
		}	
		if (isWhitespace(document.frmplan.text22.value)== false)
		{
			if (isAllNumeric(document.frmplan.text23.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Silver for range <50-100> should be numeric. \n" ;		
				}	
		}
	
	if (isWhitespace(document.frmplan.text23.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Silver for range <100 & above> is mandatory. \n" ;		
		}	
		
		if (isWhitespace(document.frmplan.text23.value)== false)
		{
			if (isAllNumeric(document.frmplan.text23.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Silver for range <100 & above> should be numeric. \n" ;		
				}	
		}
		
		//Validations for Gold plan prizing for blank fields.

		if (isWhitespace(document.frmplan.text31.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Gold for range <0-50> is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmplan.text31.value)== false)
		{
			if (isAllNumeric(document.frmplan.text31.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Gold for range <0-50> should be numeric. \n" ;		
				}	
		}
		
		if (isWhitespace(document.frmplan.text32.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Gold for range <50-100> is mandatory. \n" ;		
		}	
		
		if (isWhitespace(document.frmplan.text32.value)== false)
		{
			if (isAllNumeric(document.frmplan.text32.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Gold for range <50-100> should be numeric. \n" ;		
				}	
		}
	
		if (isWhitespace(document.frmplan.text33.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Gold for range <100 & above> is mandatory. \n" ;		
		}	
		
		if (isWhitespace(document.frmplan.text33.value)== false)
		{
			if (isAllNumeric(document.frmplan.text33.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Gold for range <100 & above>  should be numeric. \n" ;		
				}	
		}
		
		//Validations for Platinum plan prizing for blank fields.
		
		if (isWhitespace(document.frmplan.text41.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Platinum for range <0-50> is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmplan.text41.value)== false)
		{
			if (isAllNumeric(document.frmplan.text41.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Platinum for range <0-50>  should be numeric. \n" ;		
				}	
		}
		
		if (isWhitespace(document.frmplan.text42.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Platinum for range <50-100> is mandatory. \n" ;		
		}
		
		if (isWhitespace(document.frmplan.text42.value)== false)
		{
			if (isAllNumeric(document.frmplan.text42.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Platinum for range <50-100>  should be numeric. \n" ;		
				}	
		}	
	
		if (isWhitespace(document.frmplan.text43.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Platinum for range <100 & above> is mandatory. \n" ;		
		}	
		
		if (isWhitespace(document.frmplan.text43.value)== false)
		{
			if (isAllNumeric(document.frmplan.text43.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Platinum for range <100 & above>  should be numeric. \n" ;		
				}	
		}
		
		
		//Validations for Unlimited plan prizing for blank fields.
		if (isWhitespace(document.frmplan.text51.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Unlimited for range <0-50> is mandatory. \n" ;		
		}
		if (isWhitespace(document.frmplan.text51.value)== false)
		{
			if (isAllNumeric(document.frmplan.text51.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Unlimited for range <0-50>  should be numeric. \n" ;		
				}	
		}
		
		if (isWhitespace(document.frmplan.text52.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Unlimited for range <50-100> is mandatory. \n" ;		
		}	
		
		if (isWhitespace(document.frmplan.text52.value)== false)
		{
			if (isAllNumeric(document.frmplan.text52.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Unlimited for range <50-100>  should be numeric. \n" ;		
				}	
		}
	
		if (isWhitespace(document.frmplan.text53.value)== true)
		{
			ErrMsg = ErrMsg + "Price for Unlimited for range <100 & above> is mandatory. \n" ;		
		}	
		
		if (isWhitespace(document.frmplan.text53.value)== false)
		{
			if (isAllNumeric(document.frmplan.text53.value)== false)
				{
					ErrMsg = ErrMsg + "Price for Unlimited for range <100 & above>  should be numeric. \n" ;		
				}	
		}
	
		if (ErrMsg!="")
		{
		alert(ErrMsg);
		}
		
		else
		{		
			document.frmplan.action="resellers_planpricing.asp?action=save";
			document.frmplan.submit();
		}
}
</SCRIPT>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
dim sqlrange,sqlname,sqllimit,sqlid,strid,intflag
intflag=0
'Retriving the user range from the database
sqlrange = "select fld_user_range from tbl_plan_user_limit_master"
set rsrange=conn.execute(sqlrange)

if not rsrange.eof then
	strrange = trim(rsrange("fld_user_range"))
end if


'Retriving the plan user limit id depenging on the user range
sqlid = "select fld_plan_user_limit_id from tbl_plan_user_limit_master where fld_user_range='"&strrange&"'" 
set rsid=conn.execute(sqlid)

if not rsid.eof then
	strid = trim(rsid("fld_plan_user_limit_id"))
end if
	rsid.close
	rsid=nothing
'Retriving the plan name id from the database
sqlnameid = "select fld_plan_name_id from tbl_plan_name_master"
set rsnameid=conn.execute(sqlnameid)
	if not rsnameid.eof then
		strname=trim(rsnameid("fld_plan_name_id"))
	end if
	rsnameid.close
	rsnameid=nothing
'code here for showing the previous rates for easystorecreator
dim str,planid,planrate,i
redim planid(0)
redim planrate(0)
i=0

'Printing the previous rates for 0-50
str="Get_Esc_plan 1"
set strrs=conn.execute(str)

	if not strrs.eof then
		strdate=trim(strrs("fld_plan_date"))

		while not strrs.eof
			redim preserve planid(i)
			redim preserve planrate(i)
			rate = trim(strrs("fld_rate"))
			rate = formatnumber(rate,2)
			planrate(i)=rate
			planid(i)=trim(strrs("fld_plan_name_id"))
			strrs.movenext
			i=i+1
		wend
	end if
	
'Printing the previous rates for 50-100
dim str1,planid1,planrate1,j
redim planid1(0)
redim planrate1(0)
j=0
str1="Get_Esc_plan 2"
set strrs=conn.execute(str1)
	
	if not strrs.eof then
		while not strrs.eof
			redim preserve planid1(j)
			redim preserve planrate1(j)
			rate = trim(strrs("fld_rate"))
			rate = formatnumber(rate,2)
			planrate1(j)=rate
			planid1(j)=trim(strrs("fld_plan_name_id"))
			strrs.movenext
			j=j+1
		wend
	end if
	

'Printing the previous rates for 100 & above
dim str2,planid2,planrate2,k
redim planid2(0)
redim planrate2(0)
k=0
str2="Get_Esc_plan 3"
set strrs=conn.execute(str2)
	
	if not strrs.eof then
		while not strrs.eof
			redim preserve planid2(k)
			redim preserve planrate2(k)
			rate = trim(strrs("fld_rate"))
			rate = formatnumber(rate,2)
			planrate2(k)=rate
			planid2(k)=trim(strrs("fld_plan_name_id"))
			strrs.movenext
			k=k+1
		wend
	end if
	
currentdate = date()

if trim(Request.QueryString("action")) = "save" then 
	'retreive the values entered in the text boxes
		intrate11=trim(Request.Form("text11"))
		intrate12=trim(Request.Form("text12"))
		intrate13=trim(Request.Form("text13"))
		
		intrate21=trim(Request.Form("text21"))
		intrate22=trim(Request.Form("text22"))
		intrate23=trim(Request.Form("text23"))
		
		intrate31=trim(Request.Form("text31"))
		intrate32=trim(Request.Form("text32"))
		intrate33=trim(Request.Form("text33"))

		intrate41=trim(Request.Form("text41"))
		intrate42=trim(Request.Form("text42"))
		intrate43=trim(Request.Form("text43"))
		
		intrate51=trim(Request.Form("text51"))
		intrate52=trim(Request.Form("text52"))
		intrate53=trim(Request.Form("text53"))


	'code here to check if there is any value in the database
	strcheck="Get_Esc_plan 1"
	set rscheck=conn.execute(strcheck)
	if rscheck.eof then 
		'insert the values
		
		
		
		'query for inserting the Rates for Bronze
			
			conn.execute("Put_Esc_plan "&intrate11&",'"&currentdate&"',1,1")
			
			conn.execute("Put_Esc_plan "&intrate12&",'"&currentdate&"',1,2")
			
			conn.execute("Put_Esc_plan "&intrate13&",'"&currentdate&"',1,3")

	
			'query for inserting the Rates for Silver
			conn.execute("Put_Esc_plan "&intrate21&",'"&currentdate&"',2,1")
			
			conn.execute("Put_Esc_plan "&intrate22&",'"&currentdate&"',2,2")
			
			conn.execute("Put_Esc_plan "&intrate23&",'"&currentdate&"',2,3")


			'query for inserting the Rates for Gold
			conn.execute("Put_Esc_plan "&intrate31&",'"&currentdate&"',3,1")
			
			conn.execute("Put_Esc_plan "&intrate32&",'"&currentdate&"',3,2")
			
			conn.execute("Put_Esc_plan "&intrate33&",'"&currentdate&"',3,3")


			'query for inserting the Rates for Platinum
			conn.execute("Put_Esc_plan "&intrate41&",'"&currentdate&"',4,1")
			
			conn.execute("Put_Esc_plan "&intrate42&",'"&currentdate&"',4,2")
			
			conn.execute("Put_Esc_plan "&intrate43&",'"&currentdate&"',4,3")


			'query for inserting the Rates for Unlimited
			conn.execute("Put_Esc_plan "&intrate51&",'"&currentdate&"',5,1")
			
			conn.execute("Put_Esc_plan "&intrate52&",'"&currentdate&"',5,2")
			
			conn.execute("Put_Esc_plan "&intrate53&",'"&currentdate&"',5,3")

		intflag=1
	else
	 'then update the reterived values
			'query for updating the Rates for Bronze
			intrate11=formatnumber(intrate11,2)
			conn.execute("update_Esc_plan "&intrate11&",'"&currentdate&"',1,1")
			
			conn.execute("update_Esc_plan "&intrate12&",'"&currentdate&"',1,2")
			
			conn.execute("update_Esc_plan "&intrate13&",'"&currentdate&"',1,3")

	
			'query for updating the Rates for Silver
			conn.execute("update_Esc_plan "&intrate21&",'"&currentdate&"',2,1")
			
			conn.execute("update_Esc_plan "&intrate22&",'"&currentdate&"',2,2")
			
			conn.execute("update_Esc_plan "&intrate23&",'"&currentdate&"',2,3")


			'query for updating the Rates for Gold
			conn.execute("update_Esc_plan "&intrate31&",'"&currentdate&"',3,1")
			
			conn.execute("update_Esc_plan "&intrate32&",'"&currentdate&"',3,2")
			
			conn.execute("update_Esc_plan "&intrate33&",'"&currentdate&"',3,3")


			'query for updating the Rates for Platinum
			conn.execute("update_Esc_plan "&intrate41&",'"&currentdate&"',4,1")
			
			conn.execute("update_Esc_plan "&intrate42&",'"&currentdate&"',4,2")
			
			conn.execute("update_Esc_plan "&intrate43&",'"&currentdate&"',4,3")


			'query for updating the Rates for Unlimited
			conn.execute("update_Esc_plan "&intrate51&",'"&currentdate&"',5,1")
			
			conn.execute("update_Esc_plan "&intrate52&",'"&currentdate&"',5,2")
			
			conn.execute("update_Esc_plan "&intrate53&",'"&currentdate&"',5,3")

		intflag=2	
	end if
	
end if	
rscheck.close
rscheck=nothing
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
    <TD> <!--#include file="incESCmenu.asp"--></TD></TR>
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
          <form name="frmplan" method="post" >
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
              <TBODY>
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1">
          <tr>
          Plan Pricing 
          
		<hr noshade>	
          </tr>
          
          <tr bgColor=#dddddd>
			<td width="30%" align="middle">
				Plan Name
			</td>
			<% if not rsrange.eof then
					while not rsrange.eof%>
			<td width="15%" align="middle">
				<%= rsrange("fld_user_range")%>
			</td>
			
			<% rsrange.movenext
					wend
				end if %>
				<% rsrange.close
				rsrange=nothing%>
          </tr> 
          <tr>
			<td  align="middle">
				BRONZE
			</td>
		
			<td  align="middle">
				$<input value="<%= planrate(0)%>" size="6" maxlength="6" type="text" name=text11 >
			</td>
			<td align="middle">
					$<input value="<%= planrate1(0)%>"  size="6" maxlength="6" type="text" name=text12 >
			</td>
			<td align="middle">
					$<input value="<%= planrate2(0)%>" size="6" maxlength="6" type="text" name=text13 >
			</td>
          </tr>
          
           <tr>
			<td  align="middle">
				SILVER
			</td>
			<td  align="middle">
				$<input value="<%= planrate(1)%>" size="6" maxlength="6" type="text" name=text21>
			</td>
			<td align="middle">
					$<input value="<%= planrate1(1)%>" size="6" maxlength="6" type="text" name=text22 >
					
			</td>
			<td align="middle">
					$<input value="<%= planrate2(1)%>" size="6" maxlength="6" type="text" name=text23 >
			</td>
          </tr>
          <tr>
			<td  align="middle">
				GOLD
			</td>
			<td  align="middle">
				$<input value="<%= planrate(2)%>" size="6" maxlength="6" type="text" name=text31 >
			</td>
			<td align="middle">
					$<input value="<%= planrate1(2)%>" size="6" maxlength="6" type="text" name=text32 >
			</td>
			<td align="middle">
					$<input value="<%= planrate2(2)%>" size="6" maxlength="6" type="text" name=text33 >
					
			</td>
			
          </tr>
          <tr>
			<td  align="middle">
				PLATINUM
			</td>
			<td  align="middle">
				$<input value="<%= planrate(3)%>" size="6" maxlength="6" type="text" name=text41 >
			</td>
			<td align="middle">
					$<input value="<%= planrate1(3)%>" size="6" maxlength="6" type="text" name=text42 >
			</td>
			<td align="middle">
					$<input value="<%= planrate2(3)%>" size="6" maxlength="6" type="text" name=text43 >
			</td>
			
          </tr>
          <tr>
			<td  align="middle">
				UNLIMITED
			</td>
			<td  align="middle">
				$<input value="<%= planrate(4)%>" size="6" maxlength="6" type="text" name=text51 >
			</td>
			<td align="middle">
					$<input value="<%= planrate1(4)%>" size="6" maxlength="6" type="text" name=text52 >
			</td>
			<td align="middle">
					$<input value="<%= planrate2(4)%>" size="6" maxlength="6" type="text" name=text53 >
			</td>
          </tr>
           <tr>
			<td >&nbsp;
			</td>
          </tr> 
                 
          <tr>
				
				<td align="center">
					 &nbsp;&nbsp;<INPUT name="save" type="button" value="Save" onclick="javascript:fnsave()"> 
					 <input type="reset" value="Reset" name=reset >
				</td>	
				
		  </tr>
          
          </table></FORM>  
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%" style="LEFT: 0px; TOP: 178px">
             
</BODY></HTML>
<% if intflag="1" then%>
<script language=javascript>
	alert("Easystorecreators plan pricing rates have been added.")
	document.location.href = "resellers_planpricing.asp"
</script>
<% end if%>

<% if intflag="2" then%>
<script language=javascript>
	alert("Easystorecreators plan pricing rates have been updated.")
	document.location.href = "resellers_planpricing.asp"
</script>
<% end if%>
