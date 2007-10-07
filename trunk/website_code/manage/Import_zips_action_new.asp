<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Zips"
sHeader="zips"
sMenu="general"
sTitle="Import Taxes by Zipcode" 
sFullTitle = "General > <a href=tax_list.asp class=white>Taxes</a> > Import"
sMenu="general"

Dim arrColumns(5)
arrColumns(0) = "Zip_Name:Re:Text:50"
arrColumns(1) = "Zip_Start:Re:Integer:"
arrColumns(2) = "Zip_End:Re:Integer:"
arrColumns(3) = "Tax_Rate:Re:Integer:"
arrColumns(4) = "Department_Ids:Op:Text:1000"
arrColumns(5) = "Tax_Shipping:Re:Boolean:"

sProcName = "wsp_zips_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
