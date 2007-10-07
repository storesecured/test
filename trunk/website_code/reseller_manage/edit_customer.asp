<!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page displays list of the customer for the resellers site
'	Page Name:		   	resellers_customer_list.htm
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_customer_list.htm
'	Date & Time:		5 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel5=3
mSel6=2
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
<% 
dim custid
custid=Request.QueryString( "action" )
dim sqlcustomer,rscustomer,sqlgetcompname,rsdata
'code to retrive company name from store_settings table

sqlgetcompname="select store_company from store_settings left join store_logins"&_ 
			   " on store_settings.store_ID="&custid&" and store_logins.store_ID="&custid&""
				set rsdata=conn.execute(sqlgetcompname)   
				if not rsdata.eof then
					company=trim(rsdata("store_company"))
				end if
'code ends here
						
sqlcustomer="select first_name,last_name,address,city,state,zip,phone,fax,email from sys_billing where store_id="&custid
set rscustomer=conn.execute(sqlcustomer)
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
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>
          <table border="1">
          <tr>
          Customer Personal Info 
		<hr noshade>	
          </tr>
          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Company Name</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <!--<%=rscustomer("company")%>-->
          <%=company%>
          </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>First Name</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
           <%=rscustomer("first_name")%>   
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Last Name</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"> 
              <%=rscustomer("last_name")%> 
            </TD></TR>
        
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Address</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <%=rscustomer("address")%>
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>City</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <%=rscustomer("city")%>
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>State</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <%=rscustomer("state")%>
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>ZipCode</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"> 
          <%=rscustomer("zip")%>  
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Phone</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <%=rscustomer("phone")%> 
                
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Fax</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%">
          <%=rscustomer("fax")%>
            </TD></TR>
        <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2>Email</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"> 
             <%=rscustomer("email")%>   
            </TD></TR>
            
            <%
            rscustomer.close
            set rscustomer=nothing
            %>
        
        <TD noWrap width="1%">&nbsp;</TD>
        <td noWrap width="1%">
		<INPUT name="Back" type="button" value="Back" onclick="javascript:history.back()">&nbsp; 
			</td></tr>	
          </table>
          
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <TBODY>
              <TR>
                <TD height=20></TD></TR></TBODY></TABLE></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
