<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Coupon"
sHeader="coupon"
sMenu="marketing"
sTitle="Import Newsletter" 
sFullTitle = "Marketing > <a href=coupon_manager.asp class=white>Coupons</a> > Import"
sMenu="marketing"

Dim arrColumns(10)
arrColumns(0) = "Coupon_Id:Re:Text:200"
arrColumns(1) = "Coupon_Name:Re:Text:150"
arrColumns(2) = "Start_Date:Re:Date:50:"&date()
arrColumns(3) = "End_Date:Re:Date:50:1/1/2035"
arrColumns(4) = "sType:Re:Integer:"
arrColumns(5) = "Amount:Re:Integer:"
arrColumns(6) = "Total_Usage:Op:Integer::99999"
arrColumns(7) = "Customer_Usage:Op:Integer:"
arrColumns(8) = "Total_From:Op:Integer::0"
arrColumns(9) = "Total_To:Op:Integer::99999"
arrColumns(10) = "Discounted_Items_Skus:Op:Text:2000"

sProcName = "wsp_coupon_import"


%>
<!--#include file="include/import_action_include_new.asp"-->
