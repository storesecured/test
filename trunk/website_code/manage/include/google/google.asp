
 
<%

sub_write_log "in charging"
rs_store.open "Select verified_ref,grand_total from Store_Purchases WITH (NOLOCK) where oid = "&Oid&" and Store_id ="&Store_id,conn_store,1,1
If not rs_Store.eof then
    Verified_Ref=rs_Store("Verified_Ref")
	Grand_Total=rs_Store("Grand_Total")
End If
rs_Store.Close

Dim MyOrder
Set MyOrder = New Order

Dim GoogleOrderNumber, Amount, Reason, Comment
	Dim MerchantOrderNumber, Carrier, TrackingNumber, Mesage

	GoogleOrderNumber = Verified_Ref
        sub_write_log "GoogleOrderNumber=" & GoogleOrderNumber

If sType = "Capture" Then
sub_write_log "charge order"
MyOrder.ChargeOrder GoogleOrderNumber, cstr(GGrand_Total)
		 MyOrder.SendOrderCommand
elseif sType = "Credit" then

	    MyOrder.RefundOrder GoogleOrderNumber, "","", cstr(GGrand_Total)
		 MyOrder.SendOrderCommand

        trans_type=4

else
   MyOrder.CancelOrder GoogleOrderNumber,"", ""
			MyOrder.SendOrderCommand
end if

%>
