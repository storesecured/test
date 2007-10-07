<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->


<%
sql_Select_Store_Activation = "Select Invoice_header, Invoice_footer From Store_settings where Store_id = "&Store_id
rs_Store.open sql_Select_Store_Activation,conn_store,1,1
rs_Store.MoveFirst 
	Invoice_header = Rs_store("Invoice_header")
	Invoice_footer = Rs_store("Invoice_footer")
rs_Store.Close

sInstructions="The information added here will go above and below the default invoice that is created."
sArticleHelp="CustomizeLook.htm"
sFormAction = "Store_Settings.asp"
sName = "Customize_invoice"
sCommonName="Customize Invoice"
sFormName = "Customize_invoice"
sTitle = "Customize Invoice"
sFullTitle = "Design > Customize Invoice"
sSubmitName = "Update"
thisRedirect = "customize_invoice.asp"
addPicker=1
sMenu = "design"
sQuestion_Path = "design/header_and_footer.htm"
sTextHelp="templates/customize_invoice.doc"

createHead thisRedirect
%>
		
		
		<tr bgcolor='#FFFFFF'>
		<td class="inputvalue"><B>Invoice Header</b><BR>
		<%= create_editor ("Invoice_header",Invoice_header,"") %>
		<% small_help "Invoice Header" %></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		<td colspan=2>Default invoice information will be placed between your header and footer.
		</td></tr>
		<tr bgcolor='#FFFFFF'>
		<td class="inputvalue">
		<B>Invoice Footer</b><BR>
		<%= create_editor ("Invoice_footer",Invoice_footer,"") %>
		<% small_help "Invoice Footer" %></td>
		</tr>

		

<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

 }
</script>

