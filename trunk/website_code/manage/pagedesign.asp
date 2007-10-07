<%
 on error resume next

Set objHelpDict = CreateObject("Scripting.Dictionary")

show_editor_cookie=Request.Cookies("EASYSTORECREATOR")("show_editor")
if show_editor_cookie="" then
   show_editor_cookie=1
end if

if Session("Login_Privilege") <> 1 then
   Access_Priviledge = Session("Access_Priviledge")
else
   Access_Priviledge = "General,Shipping,Design,Marketing,Inventory,Statistics,Customers,Orders,My Account,Account"
end if

host = replace(lcase(request.servervariables("HTTP_HOST")),"manage.","")

sub createHead (thisRedirect)
   current_page_name=lcase(request.servervariables("SCRIPT_NAME"))
	if Session("Super_User") <> 1 then
        if Overdue_Payment > 0 then
		if current_page_name = "/old_billing_info.asp" or current_page_name = "/support_request.asp" or current_page_name = "/overdue.asp" or current_page_name = "/expired.asp" or current_page_name = "/billing.asp" or current_page_name = "/billing_info.asp" or current_page_name = "/process_billing.asp" or current_page_name = "/login_store.asp" or current_page_name = "/login_store_action.asp" or current_page_name = "/log_out.asp" or current_page_name = "/custom.asp" or current_page_name = "/custom_info.asp" or current_page_name = "/support_list.asp" or current_page_name = "/support_edit.asp" or current_page_name = "/cancel_store.asp" or current_page_name = "/my_payments.asp" or current_page_name = "/uncancel_store.asp" or current_page_name = "/my_invoice.asp" then
		else
			response.redirect "overdue.asp"
		end if
	elseif Trial_Expired=1 then
        if current_page_name = "/old_billing_info.asp" or current_page_name = "/support_request.asp" or current_page_name = "/expired.asp" or current_page_name = "/billing.asp" or current_page_name = "/billing_info.asp" or current_page_name = "/process_billing.asp" or current_page_name = "/login_store.asp" or current_page_name = "/login_store_action.asp" or current_page_name = "/log_out.asp" or current_page_name = "/custom.asp" or current_page_name = "/custom_info.asp" or current_page_name = "/support_list.asp" or current_page_name = "/support_edit.asp" or current_page_name = "/cancel_store.asp" or current_page_name = "/my_payments.asp" or current_page_name = "/uncancel_store.asp" or current_page_name = "/my_invoice.asp" then
		else
			response.redirect "expired.asp"
		end if
    end if
    end if

if Session("Login_Privilege") <> 1 then
	if instr(Access_Priviledge,"General") = 0 and sMenu="general" then
		 response.redirect "noaccess.asp?sMenu=general"
	elseif instr(Access_Priviledge,"Shipping") = 0 and sMenu="shipping" then
		 response.redirect "noaccess.asp?sMenu=shipping"
	elseif instr(Access_Priviledge,"Design") = 0 and sMenu="design" then
		 response.redirect "noaccess.asp?sMenu=design"
	elseif instr(Access_Priviledge,"Marketing") = 0 and sMenu="marketing" then
		 response.redirect "noaccess.asp?sMenu=marketing"
	elseif instr(Access_Priviledge,"Inventory") = 0 and sMenu="inventory" then
		 response.redirect "noaccess.asp?sMenu=inventory"
	elseif instr(Access_Priviledge,"Statistics") = 0 and sMenu="statistics" then
		 response.redirect "noaccess.asp?sMenu=statistics"
	elseif instr(Access_Priviledge,"Customers") = 0 and sMenu="customers" then
		 response.redirect "noaccess.asp?sMenu=customers"
	elseif instr(Access_Priviledge,"Orders") = 0 and sMenu="orders" then
		 response.redirect "noaccess.asp?sMenu=orders"
	elseif instr(Access_Priviledge,"My Account") = 0 and sMenu="account" then
		 response.redirect "noaccess.asp?sMenu=account"
	end if
end if

%>
<html>
<head>
<title><%= Name %>&nbsp;<%= sTitle %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/main.css" rel="stylesheet" type="text/css">
<script type="text/javascript" language="JavaScript1.2" src="js/apymenu.js"></script>
<script src="js/common_script.js" language="JavaScript" type="text/javascript"></script>
<% if calendar=1 then %>
	<SCRIPT LANGUAGE="JavaScript" type="text/javascript" SRC="include/CalendarPopup.js"></SCRIPT>
	<SCRIPT LANGUAGE="JavaScript" type="text/javascript">document.write(getCalendarStyles());</SCRIPT>
<% end if %>
<% if show_editor_cookie and sNeedEditor=1 then %>
<script language=JavaScript src='editor/scripts/innovaeditor.js'></script>
<% end if %>
<style>
body,html{height:100%;}
</style>
<!-- Apycom DHTML Tabs -->
<noscript><a href=http://dhtml-menu.com/>DHTML Tabs, (c)2004 Apycom</a></noscript>
<% if sNeedTabs=1 then %>


	<script type="text/javascript" language="JavaScript1.2" src="js/apytabs.js"></script>

<% end if %>
<!-- Apycom DHTML Tabs, dhtml-menu.com -->

</head>

<body bgcolor="#01324A" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"
<% if sFocus <> "" then %>
		onLoad=setTimeout("alertUser()",2400000);document.forms[0].<%= sFocus %>.focus();
	<% else %>
		onLoad=setTimeout("alertUser()",2400000);
	<% end if %>
	>

<table width="762" border="0" cellspacing="0" cellpadding="0" align="center" height="100%">
  <tr>
    <td bgcolor="#9DB9C2" width="1"><img src="images/spacer.gif" width="1" height="1"></td>
    <td width="760" valign="top" background="images/back_blue.gif"><table width="760"  border="0" cellspacing="0" cellpadding="0" bgcolor="#00517B" class="content">
        <tr>
          <td width="170" height="62" valign="top"><img src="images/top_logo.gif" width="165" height="57"></td>
          <td align=right valign=top><BR><% if sFullTitle<>"" then %><font class=white>You are here: <b><%= sFullTitle %></b></font><% end if %></td>
          <td width=20>&nbsp;</td>
        </tr>
      </table>
      <table width="760" border="0" cellspacing="0" cellpadding="0" background="">
        <tr>
          <td bgcolor="#FFFFFF"><img src="images/spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
	  <% if Store_Id<>"" then %>
      <table width="760" border="0" cellspacing="0" cellpadding="0" bgcolor="#006FA2" background="" class="content">
        <tr>
          <td width="19"><img src="images/spacer.gif" width="50" height="1"></td>
          <td width="40"><img src="images/drop_1.gif" width="1" height="38"><a href="admin_home.asp"><img src="images/menu_home.gif" width="38" height="38" border="0"></a></td>
          <td width="42"><a href="admin_home.asp" class="link2">Home</a></td>
          <td width="43"><img src="images/drop_1.gif" width="1" height="38"><a href="<%= Site_Name%>" target=_blank class="link2"><img src="images/menu_support.gif" width="42" height="38" border="0"></a></td>
		  <td width="80"><a href="<%= Site_Name%>" target=_blank class="link2">Go to site</a></td>
          <td width="44"><img src="images/drop_1.gif" width="1" height="38"><a href="livechat.asp"><img src="images/menu_help.gif" width="43" height="38" border="0"></a></td>
          <td width="65"><a href="livechat.asp" class="link2">Get Help</a></td>
          <% if show_editor_cookie=0 then
                sEditorText="Show Editor"
                sEditorLink="show_editor.asp?Status=Show"
          else
                sEditorText="Hide Editor"
                sEditorLink="show_editor.asp?Status=Hide"
          end if %>
          <td width="43"><img src="images/drop_1.gif" width="1" height="38"><a href="<%= sEditorLink %>"><img src="images/menu_resources.gif" width="42" height="38" border="0"></a></td>

          <td width="80"><a href="<%= sEditorLink %>" class="link2"><%= sEditorText %></a></td>
          <td width="43"><img src="images/drop_1.gif" width="1" height="38"><a href="log_out.asp"><img src="images/menu_logout.gif" width="42" height="38" border="0"></a></td>
          <td width="50"><a href="log_out.asp" class="link2">Logout</a></td>
          <td width="44"><img src="images/drop_1.gif" width="1" height="38"><a href="#"><img src="images/mrnu_preview.gif" width="42" height="38" border="0"></a></td>
          <td width="80"><a href=# class="link2">Site # <%=Store_Id %></a></td>
          <td width="27"><img src="images/drop_1.gif" width="1" height="38"></td>
		  <td width="19"><img src="images/spacer.gif" width="50" height="1"></td>
        </tr>
      </table>
      <table width="760" border="0" cellspacing="0" cellpadding="0" background="">
        <tr>
          <td bgcolor="#FFFFFF"><img src="images/spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="760" border="0" cellspacing="0" cellpadding="0" background="">
        <tr>
          <td bgcolor="#017FBA"><!--#include file="include/data.asp"--></td>
        </tr>
      </table>
	  <% if Session("Super_User") = 1 then %>
		  <table width="760" border="0" cellspacing="0" cellpadding="0" background="" class="content">
			<tr>
			  <td bgcolor="#017FBA"><table width="100%"><tr>
			  <td align=center>Admin Only</td>
			  <td align=center><a href=support_list_admin.asp class="link2">Support Requests</a></td>
			  <td align=center width="1%"> | </td>
			  <td align=center><a href=login_as.asp class="link2">Login As</a></td>
			  <td class="link2" align=center width="1%"> | </td>
			  <td align=center><a href=admin_settings.asp class="link2">Settings</a></td>
			  <td class="link2" align=center width="1%"> | </td>
			  <td align=center><a href=my_payments_admin.asp class="link2">Store Payments</a></td>
			  <td class="link2" align=center width="1%"> | </td>
			  <td align=center><a href=sys_template_list.asp class="link2">System Templates</a></td>
			  <td class="link2" align=center width="1%"> | </td>
			  <td align=center><a href=domain_check.asp class="link2">Check Domain</a></td>

			  
			  </tr></table></td>
			</tr>
		  </table>
		<% end if %>
	  <% end if %>
      <table width="760" border="0" cellspacing="0" cellpadding="5" background="images/back_main.gif" class="content">
        <tr>
          <td align="center"><img src="images/spacer.gif" width="8" height="8">
            <table width="750"  border="0" cellspacing="0" cellpadding="0" background="images/tbl_t.gif">
              <tr>
                <td width="54"><img src="images/hdr_<%=sMenu %>.gif" width="51" height="51"></td>
                <td width="685" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0" background="">
                    <tr>
                      <td><img src="images/spacer.gif" width="1" height="25"></td>
                    </tr>
                    <tr>
                      <td><span class="hdr"><%= sTitle %></span> </td>
                    </tr>
                  </table></td>
                <td width="12" align="right"><img src="images/tbl_tr.gif" width="11" height="51"></td>
              </tr>
            </table>
            <table width="747" border="0" cellspacing="0" cellpadding="0" background="">
              <tr>
                <td align="right" bgcolor="#FFFFFF"><table width="747" border="0" cellspacing="1" cellpadding="4" class="table" background="">
               
                    <% if 1=0 then %>
                    <tr>
                      <td align="right" bgcolor="#FFFFFF"><div align="center"></div>
                        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td nowrap><div align=center><font size=2><B><a href=usps_down.asp class=link>Updated 8:07AM PST, USPS Realtime Rates Service Restored, click here for more details</a></b></font></div></td>

                          </tr>
                        </table></td>
                    </tr>
                    <% end if %>

                    <% if Store_Cancel <>"" then %>
                    <tr>
                      <td bgcolor="#FFFF00">Your store #<%= Store_Id %> is scheduled for deletion.  If you no longer wish to have it removed please <a href=uncancel_store.asp class=link>click here</a>.  It will be removed within 48 hours.</td>
                    </tr>
                    <% end if %>
                    <tr>
                      <td bgcolor="#FFFFFF"><%= sInstructions %></td>
                    </tr>
                    <tr>
                      <td bgcolor="#FFFFFF">
                      	<form method="POST" <%= sEncType%> action=<%= sFormAction%> name=<%= sName %>>

	<input type="Hidden" name="Form_Name" value='<%= sFormName %>'>
	<input type="Hidden" name="redirect" value='<%= thisRedirect %>'>

                        <table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#D4DEE5">
                          <tr>
                            <td><table align="center" width="100%" border='0' cellpadding=3 cellspacing=1 >
<% end sub %>
<% sub createFoot (thisRedirect, thisSave)

set rs_store = Nothing
	set conn_store = Nothing

	if sSubmitName = "" then
		sSubmitName = "submit"
	end if
         %>
</table></td>
  </tr>
</table>
    <script language="JavaScript" type="text/javascript" >
		var submitcount = 0;
		function submitForm()
		{
		<%
		if thisSave=1 or thisSave=2 then %>
			if (submitcount == 0) {
				//sumbit form
				document.forms[0].<%= sSubmitName %>.value="Processing...";
				document.forms[0].<%= sSubmitName %>.disabled=true;
				submitcount ++;

				return true;
			}
			else
			{
				return false;
			}
		<% else %>
			return true;
		<% end if %>
		}
	</script>
	<% if thisSave=1 then %>
                        <table width="100%"  border="0" cellspacing="1" cellpadding="3" align="center">
                          <tr>
                            <td align="center" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="Save <%= sCommonName %>" class=buttons>
                            <% if sCancel ="" then
                                sCancel="admin_home.asp"
                            end if %>
                            <input OnClick='JavaScript:self.location="<%= sCancel %>"' name="Cancel" type="button" class="Buttons" value="Cancel">

                            </td>
                          </tr>
                        </table>
     <% end if %>
                        </td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td bgcolor="#9DB9C2" background=""><img src="images/spacer.gif" width="1" height="1"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="760" border="0" cellspacing="0" cellpadding="0" bgcolor="#00517B" class="content">
        <tr>
          <td width="746"><img src="images/spacer.gif" width="1" height="9"></td>
          <td width="14"><img src="images/spacer.gif" width="1" height="9"></td>
        </tr>
        <tr>
          <td height="20" align="right">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td align="right" height="29" valign="top"><img src="images/bottom_logo.gif" width="80" height="20"></td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td bgcolor="#9DB9C2" width="1"><img src="images/spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<% if calendar=1 then %>
	<DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
<% end if %>
<script src="js/wz_tooltip.js" language="JavaScript" type="text/javascript"></script>

</body>
</html>
<% end sub %>
<% function small_help (sJump)

	if hide_help = 3 or hide_help = 2 then
		sJump = lcase(replace(sJump,"'",""))
                sContent=objHelpDict.Item(sJump)
		if sContent="" then
                Set rs_storeHelp = Server.CreateObject("ADODB.Recordset")
		sJump = replace(sJump,"'","")
		sql_help = "select sys_help_detail.Title, sys_help_detail.Content, sys_help_detail.Link FROM sys_help WITH (NOLOCK) INNER JOIN sys_help_detail WITH (NOLOCK) ON sys_help.Topic_ID = sys_help_detail.Topic_ID WHERE (sys_help.file_name)='" & thisRedirect & "' and sys_help_detail.Title='"&sJump & "'"

		rs_storeHelp.open sql_help,conn_store,1,1
		if not rs_storeHelp.eof then
		   sContent=rs_storeHelp("Content")
		else
		   sContent=sJump
                end if
		rs_storeHelp.close
		Set rs_storeHelp = nothing
                end if
                if hide_help = 3 then %>
			</td><td class="inputvalue" align="center"><B><a class="help" onmouseover="return escape('<%= sContent %>')">(?)</a>&nbsp;&nbsp;</b>
		<% else %>
			</td><td class="inputvalue" align="center"><FONT SIZE=1><%= sContent %></FONT></b>

		<% end if

  
	else
		'response.write "</td><td class='inputvalue'>"
	end if
end function

function editor_link (sField) %>
<input class=buttons type="button" value="Editor" name="Editor" OnClick="javascript:goHtmlEditor('<%= sField %>')">
<% end function

if Session("hide_help") <> "" then
	hide_help = Session("hide_help")
elseif Request.Cookies("EASYSTORECREATOR")("hide_help") <> "" then
	hide_help = Request.Cookies("EASYSTORECREATOR")("hide_help")
else
	hide_help=3
end if
%>
