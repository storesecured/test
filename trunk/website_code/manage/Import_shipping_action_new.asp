<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Shipping"
sHeader="shipping"
sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > Budget > Import"
sImportType="Shipping"
sTitle = "Import Shipping"
sFullTitle = "Shipping > <a href=shipping_all_list.asp class=white>Methods</a> > Import"
sMenu="shipping"

Dim arrColumns(10)
arrColumns(0) = "Ship_Class:Re:Integer:"
arrColumns(1) = "Name:Re:Text:200"
arrColumns(2) = "Matrix_Low:Op:Integer::0"
arrColumns(3) = "Matrix_High:Op:Integer::0"
arrColumns(4) = "Base_fee:Op:Integer::0"
arrColumns(5) = "Weight_fee:Op:Integer::0"
arrColumns(6) = "Countries:Re:Text:4000:All Countries"
arrColumns(7) = "Ship_Location:Op:Text:200:Default"
arrColumns(8) = "Zip_Start:Op:Text:10::a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
arrColumns(9) = "Zip_End:Op:Text:10::a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
arrColumns(10) = "Always_Insert:Op:Boolean::N"

sProcName = "wsp_shipping_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
