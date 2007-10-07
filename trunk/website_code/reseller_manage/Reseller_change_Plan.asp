<!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


server.scripttimeout=1
'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the plan pricing for the reseller site.
'	Page Name:		   	reseller_change_Plan.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_change_Plan.asp
'	Date & Time:		4 Aug & 27 Aug 2004 		
'	Created By:			Sudha Ghogare
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel1=3
mSel2=2
%>

<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="../include/commonfunctions.js"></SCRIPT>
<script language="javascript">

function fnsave()
{
	var ErrMsg
	ErrMsg=""
	
	//Validation for blank fields &validation for numeric values



	

	 //if (isWhitespace(document.frmchageplan.txtu.value)== true)
	//	{
	//		ErrMsg = ErrMsg + "Monthly pricing for unlimited is mandatory. \n" ;
	//	}
	//	else if (isAllNumeric(document.frmchageplan.txtu.value)== false)
	//	{
//			ErrMsg = ErrMsg + "Monthly pricing for unlimited  should be numeric. \n" ;
//		}
//		else if ( eval(document.frmchageplan.txtEscPlan4.value) > eval(document.frmchageplan.txtu.value  ))
//		{
//			ErrMsg = ErrMsg + "Enter greater value for unlimited.\n" ;
//
//		}
	
	if (ErrMsg!="")
	{
	alert(ErrMsg);
	}
	
	else
	{
		document.frmchageplan.action="Reseller_change_plan.asp?action=save";
		document.frmchageplan.submit();
	}

}
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<%
dim intResellerID,rsGetData,sqlGetData,sqlGetData1,sqlGetData2,sqlGetData3
dim intNoCustReferred,userrange,intLimitid,intEscRates,txtgetb,txtgets
dim txtgetg,txtgetp,txtgetu,sqlGetData4,sqlGetData5,sqlGetData6,sqlGetData7,sqlGetData8

'retreving the reseller id from the session.
intResellerID = session("ResellerId")

'To find out how many customers are there for a particular reseller.
sqlGetData= "select count(store_id) as count from "&_
			" store_settings where reseller_id="&intResellerID&" and service_type>0 and trial_version=0"
set rsGetData =  conn.execute(sqlGetData)
if not rsGetData.eof then 
	intNoCustReferred=trim(rsGetData("count"))	
end if
rsGetData.close
set rsGetData=nothing

intNoCustReferred = trim(cint(intNoCustReferred))

	



'Depending on the user range find out userlimit id.
sqlGetData1= "select fld_plan_user_limit_id,fld_user_range "&_
			 " from TBL_Plan_User_Limit_Master where fld_user_low<="&intNoCustReferred&" and fld_user_high>="&intNoCustReferred
'response.write"sqlGetData1="&sqlGetData1
'response.end
set rsGetData1 =  conn.execute(sqlGetData1)


if not rsGetData1.eof then
	intLimitid= trim(rsGetData1("fld_plan_user_limit_id"))
else
   intLimitid=1
end if
rsGetData1.close
set rsGetData1=nothing

'It will show all the rates of easy store creator depending on the Userlimit id.
sqlGetData2="select fld_plan_esc_id,fld_plan_name_id,fld_plan_user_limit_id,"&_
			" fld_rate,fld_plan_date from TBL_Plan_Esc where "&_
			" fld_plan_user_limit_id="&intLimitid&" and fld_plan_name_id=5 order by fld_plan_Date"
set rsGetData2 =  conn.execute(sqlGetData2)



'Retriving the values entered by reseller from the fields
	dim intFlag 	
intFlag =0
if trim(request.querystring("action"))="save"  then 
	txtgetb=trim(Request.Form("txtb"))
	txtgets=trim(Request.Form("txts"))
	txtgetg=trim(Request.Form("txtg"))
	txtgetp=trim(Request.Form("txtp"))
	txtgetu=trim(Request.Form("txtu"))
		
	Currentdate = now()
	
	sqlCheck="select fld_rate,fld_plan_name_id from tbl_plan_user_reseller where fld_reseller_id="&intResellerID
	set rsCheck=conn.execute(sqlCheck)
	
	if not rsCheck.eof then  'then update the reterived values.
		'query for updating the Rates for Bronze.
		
		sqlupdateB=("Update_reseller_plan "&txtgetu&", '"&Currentdate&"',"&intResellerID&",5")
		conn.execute (sqlupdateB)
		
		intFlag = "1"
		
	else
		'insert the values for that particular reseller .
		

		
		sqlGetData4=("Put_reseller_plan "&intResellerID&",'"&Currentdate&"',"&txtgetu&",5 ")
		conn.execute (sqlGetData4)
		
		intFlag ="2"
	end if
	rsCheck.close
	set rsCheck=nothing
end if 

'code here for showing the previous rates for reseller
dim str,planid,planrate,j
redim planid(0)
redim planrate(0)
j=0
str="select fld_rate,fld_plan_name_id from tbl_plan_user_reseller where fld_reseller_id="&intResellerID

set strrs=conn.execute(str)
if not strrs.eof then
		while not strrs.eof
			redim preserve planid(j)
			redim preserve planrate(j)
			planrate(j)=trim(strrs("fld_rate"))
			test = split(planrate(0),".")
			if test(1) ="" then
				planrate(j) = planrate(j)&".00"
			end if
			planid(j)=trim(strrs("fld_plan_name_id"))
			strrs.movenext
			j = j +1
		wend
end if	
strrs.close
set strrs=nothing

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
               href="Reseller_change_Logo.asp">Change Logo 
              </A></TD>
        </TR>
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="Reseller_change_HKeywords.asp">Manage Keywords 
 
              </A></TD>
        </TR>      
        
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
                href="Reseller_change_Plan.asp">Change Plan Pricing 
  
              </A></TD>
        </TR>  
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
                href="Reseller_change_Desc.asp">Define Description               </A></TD>
        </TR>                 
 <%				
'******************************Code Added Here:14th JAN 2005 ******************************
%>
        <TR>
            <TD class=meniu height=20 
							onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
							onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
							<A class=b href="Reseller_display_name.asp">Choose Name</A>
						</TD>
        </TR>
				<TR>
            <TD class=meniu height=20 
							onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
							onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
							<A class=b href="Reseller_contact_us.asp">Contact us</A>
						</TD>
        </TR>
<%
'*****************************************************************************
%>        
				</TBODY>
			</TABLE>
		</TD>
    <TD height=15 vAlign=top width=570></TD>
	</TR>

        <TR>
          
          <TD class=pagetitle height=400 vAlign=top><FONT 
            color=red>
          
          Note: <font face="arial" size="2">Your customers will be offered the option to receive a 5% discount by paying quarterly, 10% discount by paying semi-annually, and 25% discount by paying yearly. </font><!--<A href="javascript:fnopen()" title="Click here to know pricing plan">Easystorecreator Plan Pricing</A>-->
          <br>
          <br></FONT>
          Change Plan Pricing 
          <form action="" method="post" name="frmchageplan">
          <TABLE align=left border=1 cellPadding=3 cellSpacing=0 width=525 id=TABLE1 >
        <TBODY>
        <br>
        <TR bgColor=#dddddd>
        
          <TD width="50%"><B>Plan Pricing</B></TD>
          <TD align=middle width="10%"><B>STORE</B></TD></TR>
        <TR>
          <TD>Monthly Pricing</TD>
<input type=hidden size="7" maxlength="7" name="txtb" value=<%=FormatNumber(planrate(0),2)%> ></TD>
<input type=hidden size="7" maxlength="7" name="txts" value=<%=FormatNumber( planrate(1),2)%> ></TD>
<input type=hidden size="7" maxlength="6" name="txtg" value=<%=FormatNumber(planrate(2),2)%> ></TD>
<input type=hidden size="7" maxlength="6" name="txtp" value=<%=FormatNumber(planrate(3),2)%> ></TD>
          <TD >$<input size="7" maxlength="7" name="txtu" value=<%=FormatNumber(planrate(4),2)%> ></TD></TR>
          <TR>
          <TD>Easystorecreator Buy Rates </TD>
          
          <%
          dim i
          i=0
          if not rsGetData2.eof then
			while not rsGetData2.eof and i < 10
				intEscRates= rsGetData2("fld_rate")
				intPlanId= rsGetData2("fld_plan_name_id")
				
				'To retrieve the name of the plan
				sqlGetData3= "select fld_plan_name_id,fld_plan_name from TBL_Plan_name_master where 				fld_plan_name_id="&intPlanId

					
				set rsGetData3 =  conn.execute(sqlGetData3)
				if not rsGetData3.eof then
					strPlanNAme = rsGetData3("fld_plan_name")
				end if
				rsGetData3.close
				set rsGetData3=nothing
				name = "txtEscPlan"&i
				
            %>
          <TD >$ <%= formatnumber(intEscRates,2)%></TD>
          <input type ="hidden" value="<%= intEscRates%>" name="<%=name%>">
          <%
          i = i +1
          rsGetData2.movenext
			wend
		end if
			rsGetData2.close
			set rsGetData2=nothing
		%>
          </TR>
			<tr>
		<tr>
		<td>
					 <INPUT name="save" type="button" value="Save" onclick="javascript:fnsave()">&nbsp; 
<INPUT id=reset name=reset type=reset value=Reset>
				</td>	
        </tr>      
		
        
</TBODY> </TABLE>
<table>

</table>			  
				</TD>
				<td>
				</td></TR>	
              
              <TBODY>
        <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>

<% if intFlag ="1" then%>
<script language="javascript">
	alert("The plan rates have been updated.")
	document.location.href = "Reseller_change_Plan.asp"
</script>
<%end if%>

<% if intFlag ="2" then%>
<script language="javascript">
	alert("The plan rates have been added.")
	document.location.href = "Reseller_change_Plan.asp"
</script>
<%end if%>
