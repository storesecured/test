<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

'RETRIVE CURRENT VALUES FROM THE DATABASE

sTitle = "Design Changes"
thisRedirect = "designmoved.asp"
sMenu = "design"
createHead thisRedirect
%>

			 <tr><TD>The page you are trying to access has been moved as part of a design interface upgrade.  
                         To view the <a href=designfaq.asp class=link>faq</a> on this upgrade please <a href=designfaq.asp class=link>click here</a>.<BR><BR>
                         The information you were looking for can now be found under <a href=template_list.asp class=link>Design-->Custom Templates</a> or by <a href=template_list.asp class=link>clicking here</a></TD></tr>


		
				<% createFoot thisRedirect,0 %>
