<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page shows the pages for changing the description for reseller site.
'	Page Name:		   	reseller_change_desc.asp
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_change_desc.asp
'	Date & Time:		4 Aug 2004 		
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
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language="javascript">


</script>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">

<%
'select query for the pages
dim rsGetData,sqlGetData

sqlGetData= "select fld_page_id,fld_page_name from TBL_Page_master order by fld_page_name asc"
set rsGetData=conn.execute(sqlGetData)
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
                href="Reseller_change_Desc.asp">Define Description</A></TD>
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
							<A class=b href="Reseller_display_name.asp">Contact us</A>
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
          <TD class=pagetitle height=400 vAlign=top>Define Description
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action="" method=post name="frmchangedesc">
              <table border="1" width=100%>
              <br>
              <tr>
              <td>
				<b>Page Name</b>
				
              </td>
				
              <td>
              
              </td>
              
              </tr>
              <tr>
				  <td>
				  &nbsp;
				  </td>
			  <td>
              
              </td>
              
              </tr>
              
            <%
            if not rsGetData.eof then 
				while not rsGetData.eof 
            %>
              <tr>
              <td>
				<%=rsGetData("fld_page_name")%>
              </td>
			
              <td>
             	<a href="Reseller_page_desc.asp?action=<%=rsGetData("fld_page_id")%>">Edit</a>
              </td>
              </tr>
              <% 
				rsGetData.movenext
				wend
				             
             end if%>
              <% rsGetData.close
              rsGetData=nothing%>
              </table>
              
               </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
