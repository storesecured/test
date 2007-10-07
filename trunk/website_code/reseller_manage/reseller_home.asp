<%
if trim(Request.Form("hidsession"))<> "" then 
	session("ResellerID") = trim(Request.Form("hidsession"))
	
end if


if trim(Request.Form("txtUserId"))<> "" then 
	'taking values entered by user for username and password in variable. 
	txtid=trim(Request.Form("txtUserId"))
	txtpass=trim(Request.Form("txtPassword"))
	set conn1 = server.CreateObject("adodb.connection")
	strconn1 = "DRIVER=SQL Server;SERVER=10.235.158.138;UID=melanie;PWD=tom237;DATABASE=wizard"		
	'strconn1 = "DRIVER=SQL Server;SERVER=agni;UID=user;PWD=thinkmore;DATABASE=esc"		
	conn1.open strconn1
	'query for retrieving database values
		sqlgetdata = "select fld_reseller_id,fld_user_name,fld_password from tbl_reseller_master where fld_user_name='"&txtid&"' and "&_
					"fld_password='"&txtpass&"'"
		set rscheck = conn1.execute(sqlgetdata)


		if not rscheck.eof then 

			intResellerID = trim(rscheck("fld_reseller_id"))
			session("ResellerID")  = intResellerID
		else
			
			Response.Redirect "Reseller_Error.asp?error=3"
			Response.End
				
		end if
end if		

%>	
<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This the reseller's home page
'	Page Name:		   	reseller_home.asp
'	Version Information:EasystoreCreator
'	Input Page:		    it is the home page for reseller's login
'	Output Page:	   
'	Date & Time:		9th August
'	Created By:			Devki Anote
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 

mSel1 = 3
mSel2 = 2
%>

<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
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
     <!--#include file="incmenu.asp" -->       
            
            
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
          <TD class=pagetitle height=400 vAlign=top>Through this section you can customize the sales website.
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%"><FORM action=Page_action.asp method=post 
              name=""><INPUT name=Form_Name type=hidden value=Create_Page> 
              <INPUT name=redirect type=hidden value=new_page.asp> 
              <TBODY>
              <TR>
         

                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
