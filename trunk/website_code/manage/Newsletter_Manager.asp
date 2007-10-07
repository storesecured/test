<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "marketing/newsletter.htm"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Newsletter"
myStructure("ColumnList") = "newsletter_id,email_address,first_name,last_name"
myStructure("DefaultSort") = "email_address"
myStructure("PrimaryKey") = "newsletter_id"
myStructure("Level") = 1
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "newsletter_manager.asp"
myStructure("FormAction") = "newsletter_manager.asp"
myStructure("Title") = "Newsletter Subscribers"
myStructure("FullTitle") = "Marketing > Newsletter > Subscribers"
myStructure("CommonName") = "Newsletter Subscriber"
myStructure("NewRecord") = "new_newsletter.asp"
myStructure("Heading:newsletter_id") = "PK"
myStructure("Heading:email_address") = "Email"
myStructure("Heading:first_name") = "First"
myStructure("Heading:last_name") = "Last"
myStructure("Format:email_address") = "STRING"
myStructure("Format:first_name") = "STRING"
myStructure("Format:last_name") = "STRING"
%>
<!--#include file="head_view.asp"-->
<tr bgcolor='#FFFFFF'>
	<td width="100%" colspan="3" height="15">&nbsp;
	<input type="button" OnClick=JavaScript:self.location="message_customers.asp" class="Buttons" value="Send Newsletter" name="Create_new_Page_link">
	</td>
</tr>
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then

	end if
	if Request.Form("Delete_Id") ="SEL" then
  	if request.form("DELETE_IDS") <> "" then

	  end if
  end if
createFoot thisRedirect, 0
%>

