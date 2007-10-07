<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Secure Certificate"
thisRedirect = "cert.asp"
sMenu="none"
createHead thisRedirect  %>



<TR bgcolor='#FFFFFF'><td><B>New secure graphic</td></tr>
<TR bgcolor='#FFFFFF'><TD>The old secure graphic shown for users who have it enabled in their stores has been replaced.  A view of what the new certificate looks like is shown below.  To turn this graphic on/off from your credit card entry page please visit Advanced-->Misc Settings on the Checkout tab check or uncheck the option to show the secure logo.<BR><!-- Secure ID - Do Not Remove! --><div id="digicertsitesealcode" style="width: 81px; margin: 10px auto 10px 10px;" align="center"><script language="javascript" type="text/javascript" src="https://www.digicert.com/custsupport/sealtable.php?order_id=<%=ssl_oid%>&seal_type=a&seal_size=large"></script><script language="javascript" type="text/javascript">coderz();</script></div><!-- END Secure ID -->
</textarea></td></tr>


<% createFoot thisRedirect, 0 %>


