<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Newsletter"
sHeader="newsletter"
sMenu="marketing"
sTitle="Import Newsletter" 
sFullTitle = "Marketing > <a href=newsletter_manager.asp class=white>Newsletter</a> > Subscribers > Import"
sMenu="marketing"

Dim arrColumns(2)
arrColumns(0) = "Email:Re:Text:50"
arrColumns(1) = "First_Name:Op:Text:50"
arrColumns(2) = "Last_Name:Op:Text:50"
sProcName = "wsp_newsletter_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
