<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Activate"
thisRedirect = "expired.asp"
sMenu="account"
createHead thisRedirect
 %>



		<TR bgcolor='#FFFFFF'><td>

		You will not be
		able to manage your store until you have activated it.
      <BR><BR>All activations are backed by
		our no risk, 45 day money back guarantee.  If you are unhappy for any reason simply cancel 
		your account and request a refund.
		</td></tr>

                <TR bgcolor='#FFFFFF'><TD>
		<BR><a href=billing.asp class=link><B>Click here to activate now.</b></a>
		</td></tr>

		<TR bgcolor='#FFFFFF'><TD>
		<a href=cancel_store.asp class=link><BR><BR>Click here to cancel your store</a>
		</td></tr>
		<TR bgcolor='#FFFFFF'><TD>
		<a href=support_list.asp class=link><BR><BR>Click here to contact support</a>
		</td></tr>

	  <TR bgcolor='#FFFFFF'><TD></td></tr>

<% createFoot thisRedirect, 0 %>


