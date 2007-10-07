<!-- #include file="functions.asp" -->

<%
sub create_form_post_protx ()
  If Auth_Capture then
	  x_type = "PAYMENT"
  else
	  x_type = "DEFERRED"
  end if
  if Store_Id=0 then
   response.write "<form method='POST' action='https://ukvpstest.protx.com/vps2form/submit.asp' name='Payment'>"
  else
	response.write "<form method='POST' action='https://ukvps.protx.com/vps2form/submit.asp' name='Payment'>"

  end if
	response.write "<input type='hidden' name=VPSProtocol value='2.22' onSubmit=""return checkFields();""><input type='hidden' name=TxType value='"&x_type&"'>"

end sub

sub create_form_content_protx ()
	sql_real_time = "exec wsp_real_time_property "&Store_Id&","&Real_Time_Processor&";"
    rs_Store.open sql_real_time,conn_store,1,1
	  rs_Store.MoveFirst
	  Do While Not Rs_Store.EOF
		  select case Rs_store("Property")
		case "VendorName"
		  VendorName = decrypt(Rs_store("Value"))
		case "VendorPassword"
			VendorPassword = decrypt(Rs_store("Value"))
		case "VendorCurrency"
		  VendorCurrency = decrypt(Rs_store("Value"))
		end select
	 Rs_store.MoveNext
	Loop
	Rs_store.Close
	if VendorName="" or VendorPassword="" then
		fn_error "The process can't be done ,Please contact store owner for more information"

	else
	ii=now()
        h=hour(ii)
        m=minute(ii)
        s=second(ii)
        d=day(ii)
        mm=month(ii)
        y=year(ii)
        kk=d&mm&y&h&m&s
        strTxt = "VendorTxCode=" & oid & "A"&kk & "&"
	strTxt = strTxt + "Amount=" & GGrand_Total & "&"
	strTxt = strTxt + "Currency=" & VendorCurrency & "&"
	strTxt = strTxt + "Description=Invoice"& Oid & "&"
	strTxt = strTxt + "SuccessURL=" & Site_Name&"include/protx/protx.asp?Shopper_Id="&Shopper_id&"&"
	strTxt = strTxt + "FailureURL=" & Site_Name&"include/protx/protx.asp?Shopper_id="&Shopper_id&"&"
	strTxt = strTxt + "VendorEMail=" & Store_Email & "&"
	strTxt = strTxt + "EMailMessage=" & "&"

	rs_store.open "select * from store_customers where record_type=0 and cid="&cid&" and store_id="&Store_Id, conn_store, 1, 1
	If not rs_Store.eof then

		strTxt = strTxt + "CustomerEMail=" & rs_store("Email") & "&"
		strTxt = strTxt + "CustomerName=" & rs_store("First_name") & " " & rs_store("Last_name") & "&"
		strTxt = strTxt + "BillingAddress=" & rs_store("Address1") & "&"
		strTxt = strTxt + "BillingPostCode=" & rs_store("Zip") & "&"

	End If
	rs_store.close

	rs_store.open "Select * from Store_Purchases where Shopper_id ="&Shopper_ID&" AND oid = "&Oid&" and Store_id ="&Store_id, conn_store, 1, 1
	If not rs_Store.eof then
	
		strTxt = strTxt + "DeliveryAddress=" & rs_store("ShipAddress1") & "&"
		strTxt = strTxt + "DeliveryPostCode=" & rs_store("ShipZip")

	End If
	rs_store.close
	crypt = base64Encode(SimpleXor(strTxt,VendorPassword))
	
	response.write "<input type='hidden' name=Crypt value='"&Crypt&"'>"
	response.write "<input type='hidden' name=Vendor value='"&VendorName&"'>"
	end if
end sub

%>

