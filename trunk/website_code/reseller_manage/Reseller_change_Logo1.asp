<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page uploads the logo for the resellers site
'	Page Name:		   	reseller_change_logo.htm
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_change_logo.htm
'	Date & Time:		4 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 

mSel1 = 3
mSel2 = 2
%>


<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
<LINK href="images/style.css" rel=stylesheet type=text/css>
<script language="javascript" src="../include/commonfunctions.js"></script>


<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
'ASP code here
dim intResellerID,strlogotitle,sqlPutData,strlogoimage,sqlLogo,rsLogo,NewFileName,strFileName,objUpload,strPath
dim TDateu,strhidimage,intflag
intResellerID = session("ResellerID") 
'intResellerID = 1

'code here to check whether the file is already there or not into the database whenever the page is visited
sqlLogo= "select fld_logo_title,fld_title_image from tbl_Reseller_Logo where fld_reseller_id="&intResellerID 
set rslogo=conn.execute(sqlLogo)
if not rslogo.eof then 
	imagename = trim(rslogo("fld_title_image"))
	oldimagename =imagename
	strlogotitle =trim(rslogo("fld_logo_title"))
end if
rslogo.close
set rslogo=nothing



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
        
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Change Logo
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              
              <FORM action="reseller_change_logo_action.asp" method=post name="frmChangeLogo" encType="multipart/form-data">
              <TBODY>
              
				<%
				if imagename<> "" then 
				
				imagename=  right(imagename,len(imagename)-instr(1,imagename,"_"))
				else
					imagename = "Not Uploaded"
				end if
			
				%>
								<tr>
				
            	 <td>
            	 <br>
					Existing Image 
				</td>
				<td>
				
					 <%=imagename%>
					
					
				</td>	
              <br>
              <tr>
                
				<tr>
            	 <td>
					  Logo Image
				</td>
				<td>
					 <INPUT name="File1" type="file"> 
				</td>	
              <br>
              <tr>
                
                <td>
					  
				</td>
				<td>
					 <INPUT name="save" type="submit" value="Save"> 
					 <input type="reset" value="Reset" >
				</td></tr>	
				<input type="hidden" name="hidimage" value="<%=oldimagename%>">
              
              <TBODY>
              <TR>
				
                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE>
              
              </FORM>
</BODY></HTML>
<%
if intFlag="1" then
%>
<script language="javascript">
alert("Logo image has been added.")
document.location.href="reseller_change_logo.asp"
</script>
<%
end if
%>
<%
if intFlag="2" then
%>
<script language="javascript">
alert("Logo image has been updated.")
document.location.href="reseller_change_logo.asp"
</script>
<%
end if
%>
