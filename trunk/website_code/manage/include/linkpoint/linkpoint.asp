<%

'*****************************************************************************
' Copyright 2003 LinkPoint International, Inc. All Rights Reserved.
' 
' This software is the proprietary information of LinkPoint International, Inc.  
' Use is subject to license terms.
'
'******************************************************************************    
sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"

rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
  Do While Not Rs_Store.EOF
		select case Rs_store("Property")
	  case "storename"
		  storename = decrypt(Rs_store("Value"))
		end select
		Rs_store.MoveNext
  Loop
Rs_store.Close

C_Configfile = storename
Key_Folder = fn_get_sites_folder(Store_Id,"Key")
C_Keyfile	 = Key_Folder&"cert.pem"

Const C_Host  = "secure.linkpt.net"
Const C_Port  = 1129

rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1
If not rs_Store.eof then
	first_name=rs_Store("First_name")
	last_name=rs_Store("Last_name")
	address1=rs_Store("Address1")
	address2=rs_Store("Address2")
	city=rs_Store("City")
	zip=rs_Store("zip")
	country=rs_Store("Country")
	phone=rs_Store("Phone")
	fax=rs_Store("fax")
	ccid = rs_Store("CCID")
	if rs_Store("Tax_Exempt") then
		TaxExempt="Y"
	else
		 TaxExempt="N"
	end if
End If
rs_store.close

rs_store.open "Select * from Store_Purchases where oid = "&Oid&" and Store_id ="&Store_id,conn_store,1,1
If not rs_Store.eof then
	Shipping_Method_Price = Rs_store("Shipping_Method_Price")
	Tax = Rs_store("Tax")
	Cust_PO = Rs_store("Cust_PO")
	ShipFirstname=rs_Store("ShipFirstname")
	ShipLastname=rs_Store("ShipLastname")
	ShipAddress1=rs_Store("ShipAddress1")
	ShipAddress2=rs_Store("ShipAddress2")
	ShipCity=rs_Store("ShipCity")
	Shipzip=rs_Store("Shipzip")
	Shipstate=rs_Store("Shipstate")
	ShipCountry=rs_Store("ShipCountry")
	ShipPhone=rs_Store("ShipPhone")
	ShipFax=rs_Store("ShipFax")
	ShipEmail=rs_Store("ShipEmail")
	ShipCompany=rs_Store("ShipCompany")
	Verified_Ref=rs_Store("Verified_Ref")
	CardNumber=rs_Store("CardNumber")
	CardExpiration=rs_Store("CardExpiration")
	AuthNumber=rs_Store("AuthNumber")
End If
rs_Store.Close

rs_store.open "select country_code from sys_countries where country='"&country&"'"
if not rs_store.eof then
	country_code=rs_Store("Country_Code")
end if
rs_store.close

rs_store.open "select country_code from sys_countries where country='"&ShipCountry&"'"
if not rs_store.eof then
	shipcountry_code=rs_Store("Country_Code")
end if
rs_Store.Close

Dim IsPostBack



 ' process order
  ProcessOrder()
  ' redirect to response page

  ParseResponse(Session("resp"))

  Verified_Ref = Session("Verified_Ref")
AuthNumber = Session("AuthNumber")
avs_result = Session("avs_result")




%>

