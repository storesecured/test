<!--#include file="include/header.asp"-->

<TABLE width="100%">


    <% if Enable_affiliates <> 0 then %>
	 <TR><TD><b><a href="<%= Site_Name %>affiliate_program.asp" class='link'>Affiliate Program</a></b></TD></TR>
    <% end if %>
	 <TR><TD><b><a href="<%= Site_Name %>Contactus.asp" class='link'>Contact Us</a></b></TD></TR>
    <TR><TD><b><a href="<%= Site_Name %>privacy.asp" class='link'>Privacy Policy</a></b></TD></TR>
    <TR><TD><b><a href="<%= Site_Name %>returns.asp" class='link'>Return Policy</a></b></TD></TR>


</TABLE>

<!--#include file="include/footer.asp"-->
