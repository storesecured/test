<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Hacker free"
thisRedirect = "hackerfree.asp"
sMenu="none"
createHead thisRedirect  %>


      
		<tr bgcolor='FFFFFF'><td><B>Hacker Free Certification and Free PCI Pass</td></tr>
		<tr bgcolor='FFFFFF'><TD>StoreSecured has partnered with Xentinel Security to offer you
		free PCI compliance and the ability to receive hacker free certification
		seals for your ecommerce websites.
		<BR><BR>
		<a href=https://secure.xentinelsecurity.com/resellers/enroll.aspx?ID=8992 class=link target=_blank>PCI Pass</a> is a free PCI compliance certification for all StoreSecured merchants
		<BR><BR>
		<a href=https://secure.xentinelsecurity.com/resellers/enroll.aspx?ID=8992 class=link target=_blank>Hacker Free certification</a> reassures your customers that your site is Hacker Free.
                Xentinel security runs a full scan of your site each and everyday to verify that 
                it is in fact free of vulnerabilities.  The Hacker Free seal can be added to your site to give your
		customers additional trust that your site is secure for as little as $17.99 per month.  Hacker
		free certification has shown to increase confidence and sales by as much as 15%.
		<BR><BR>
		Below is an example of our hacker free security tag for the site you are currently on
		<BR>
		<!-- START OF HACKER FREE SECURITY TAG -->
<a target="_blank" href="https://secure.xentinelsecurity.com/Certificates/Verify.aspx?ref=<%= hackerfree_url %>">
<img width =139 height =40 src="//secure.xentinelsecurity.com/Certificates/sealserver.aspx?ref=<%= hackerfree_url %>&Seal=1"
 OnError = "this.onerror=null;this.src='';this.width=0;this.height=0"
 border = "0" name = "hs_seal" alt="HACKER FREE certified websites are a secure place to shop, they are 99.9% immune to hacker attacks."></a>
<!-- END OF HACKER FREE SECURITY TAG -->
<BR><BR>
<a href=https://secure.xentinelsecurity.com/resellers/enroll.aspx?ID=8992 class=link target=_blank>Click here to find out more or signup for your own PCI Pass or Hacker Free seal today!</a>
                </td></tr>


<% createFoot thisRedirect, 0 %>


