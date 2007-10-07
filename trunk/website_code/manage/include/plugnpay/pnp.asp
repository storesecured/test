<%

sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF
	select case Rs_store("Property")
		case "publisher-name"
			publisher_name = decrypt(Rs_store("Value"))
	end select
	Rs_store.MoveNext
Loop
Rs_store.Close

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
	state=rs_Store("State")
	ccid = rs_Store("CCID")
	if rs_Store("Tax_Exempt") then
		TaxExempt="TRUE"
	else
		 TaxExempt="FALSE"
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

GGrand_Total = FormatNumber(GGrand_Total,2)
if sType="Capture" then
   sTypeTrans = "mark"
elseif sType = "Credit" then
   sTypeTrans = "credit"
else
    sTypeTrans = "void"
end if


	 Set PNPObj = CreateObject("pnpcom.main")

	 STR = "orderID="&Verified_Ref
	 STR = STR	& "&mode="&sTypeTrans
	 STR = STR	& "&txn-type=auth"
	 STR = STR	& "&publisher-name="&publisher_name
	 STR = STR	& "&publisher-password=pnpdemo"
	 STR = STR 	& "&card-amount="&GGrand_Total




 Results = PNPObj.doTransaction("",STR,"","")

 Dim objResults
 Set objResults = Server.CreateObject("Scripting.Dictionary")

 myArray = split (Results,"&")
 for i = 0 to UBound(myArray)
   myArray(i) = replace(myArray(i),"+"," ")
   myArray(i) = replace(myArray(i),"%20"," ")
   myArray(i) = replace(myArray(i),"%21","!")
   myArray(i) = replace(myArray(i),"%23","#")
   myArray(i) = replace(myArray(i),"%24","$")
   myArray(i) = replace(myArray(i),"%25","%")
   myArray(i) = replace(myArray(i),"%26","&")
   myArray(i) = replace(myArray(i),"%27","'")
   myArray(i) = replace(myArray(i),"%28","(")
   myArray(i) = replace(myArray(i),"%29",")")
   myArray(i) = replace(myArray(i),"%2c",",")
   myArray(i) = replace(myArray(i),"%2d","-")
   myArray(i) = replace(myArray(i),"%2e",".")
   myArray(i) = replace(myArray(i),"%40","@")
   pos = inStr(1,myArray(i),"=")
   if (pos > 1) then
     myKey = Left(myArray(i),pos-1)
     myVal = Mid(myArray(i),pos+1)
     response.write myKey + " = " + myVal + "<br>"
     objResults.Item(myKey) = myVal
   End If
 Next
      
 if(objResults.Item("FinalStatus") = "failure" or objResults.Item("FinalStatus") = "problem") then
   response.redirect "error.asp?Message_id=101&Message_Add="&Server.UrlEncode(objResults.Item("MErrMsg"))
 End If
 

%>
