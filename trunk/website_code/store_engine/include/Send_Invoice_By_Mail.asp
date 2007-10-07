<!--#include file="cart_display.asp"-->
<%

'retrieving cutomer info for the plain text invoice
Function Get_customer_info_text(cid,Record_type)
	sql_select_cust = "exec wsp_customer_lookup "&store_id&","&cid&","&Record_type&";"
    fn_print_debug sql_select_cust
    rs_Store.open sql_select_cust,conn_store,1,1
	rs_Store.MoveFirst
	    return_info  = rs_Store("Company") &  vbcrlf & rs_Store("First_name") & chr(32) & rs_Store("Last_name") & vbcrlf & rs_Store("Address1") & vbcrlf & rs_Store("Address2") & vbcrlf & rs_Store("City") & ", " & rs_Store("State") & ", " & rs_Store("zip") & vbcrlf & rs_Store("Country") & vbcrlf &"Phone: " & rs_Store("Phone") & vbcrlf &"Fax: " & rs_Store("Fax") & vbcrlf & rs_Store("EMail") 
	rs_Store.Close
	Get_customer_info_text = return_info
End Function

'CREATE AND SEND AN INVOICE BY MAIL
sub Send_Invoice_By_Mail(oid,Shopper_Id,sentfrom,sendto,Subject,Body)
    
    'CREATE A TEMPORARY FILE IN EXPORT FOLDER
	Set objFileSys = Server.CreateObject("Scripting.FileSystemObject")
	Export_Folder = fn_get_sites_folder(Store_Id,"Export")  
	
	if not objFileSys.FolderExists(Export_Folder) then
		Set f1 = objFileSys.CreateFolder(Export_Folder)
	end if
	varOutputFile = Export_Folder& "invoice_" & oid & ".htm"
	on error goto 0
	Set objInvoiceOut = objFileSys.CreateTextFile(varOutputFile, True)

	'WRITE THE HTML CONTENTS OF THE INVOICE INTO THE TEMPORARY FILE
    objInvoiceOut.WriteLine("<html>")
    objInvoiceOut.WriteLine("<head>")
    objInvoiceOut.WriteLine("<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">")
    objInvoiceOut.WriteLine("<title>Customer Receipt</title>")
    objInvoiceOut.WriteLine("<style type='text/css'>body,p,table,td,ul,ol")
    objInvoiceOut.WriteLine("{font-family:verdana,helvetica,arial,sans-serif;")
    objInvoiceOut.WriteLine("font-size:10pt}")
    objInvoiceOut.WriteLine("</style>")
    objInvoiceOut.WriteLine("<base href='"&Site_Name&"'>")
    objInvoiceOut.WriteLine("</head>")
    objInvoiceOut.WriteLine("<body>")
    objInvoiceOut.WriteLine("<table width=""600""><tr><td>")
    objInvoiceOut.WriteLine(fn_display_invoice (oid,shopper_id,1,1))
    objInvoiceOut.WriteLine("</td></tr></table>")
    objInvoiceOut.WriteLine("</body></html>")
	objInvoiceOut.Close

    if Html_Notifications or isNull(Html_Notifications) then
        isHtml=True
        Body="<html><head><base href='"&Site_name&"'>"&_
	   	"<style type='text/css'>.normal"&_
    		"{font-family:verdana,helvetica,arial,sans-serif;"&_
    		"font-size:10pt}"&_
		"</style></head><body>"&Body&"</body></html>"
    else
        isHtml=False
    end if

	for each one_SendTo in Split(sendto, ",")
        if one_SendTo <> "" then
            Set Mail = Server.CreateObject("Persits.MailSender")
            Mail.From = SentFrom
            fn_print_debug "from="&sentfrom
            Mail.AddAddress one_SendTo
            fn_print_debug "to="&one_sendto
            Mail.Subject = Subject
            Mail.Body = Body
            Mail.AddAttachment varOutputFile
            Mail.IsHTML = isHtml
            Mail.Queue = True
		    Mail.Send
        end if
    next
	'DELETE THE TEMPORARY FILE
    objFileSys.DeleteFile(varOutputFile)
end sub
                     
'sending invoice as plain text
Function Send_Text_Invoice_By_Mail(oid)
	sPrintInvoiceText = fn_display_invoice (oid,shopper_id,1,1)
	if Html_Notifications then
	else
		sPrintInvoiceText=replace(sPrintInvoiceText,"<HR>",vbcrlf&"------------------------------------------------------------------------------------------"&vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"&nbsp;"," ")
		sPrintInvoiceText=replace(sPrintInvoiceText,"</td>",chr(9))
		sPrintInvoiceText=replace(sPrintInvoiceText,"</TD>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"</tr>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"</TR>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"<br>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"<BR>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"<blockquote>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"</blockquote>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"<div>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"</div>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"<p>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,"</p>",vbcrlf)
		sPrintInvoiceText=replace(sPrintInvoiceText,sLineBreak,vbcrlf)
		sPrintInvoiceText=fn_stripHtml(sPrintInvoiceText)
	end if
	Send_Text_Invoice_By_Mail=sPrintInvoiceText
End Function

sub send_order_notifications(loid)
    fn_print_debug "in send order notification "&loid
    if order_notification_to_admin_enable then
        fn_print_debug "in order_notification_to_admin_enable"
	    'SEND INDEPENDENT EMAIL TO STORE ADMIN
	    Subject = "StoreSecured new order notification: "&loid&" for "&store_name
	    Body = "Store Owner"&chr(10)&"You have received a new order #"&loid&" at "&FormatDateTime(Now(),2)&chr(10)&" To process and view the details of this order please click here http://manage.storesecured.com/order_details.asp?oid="&loid&"&cid="&cid&chr(10)
	    Call Send_Mail(sNoReply_email,Store_Email,Subject,Body)
	    send_to = email&","&store_email&","&fn_replace(order_notification_sent_to, " ", "")
    else
        fn_print_debug "in else order_notification_to_admin_enable"
	    send_to = email&","&order_notification_sent_to 
    end if
    if order_notification_enable then
        fn_print_debug "in order_notification_enable"
	    'IF SET TO SEND MAIL TO STORE ADMIN ALSO

        order_notification_body_replaced=order_notification_body
	    if InStr(order_notification_body,"%") > -1 then
            order_notification_body_replaced = Replace(order_notification_body_replaced,"%ORDER%",loid)
            order_notification_body_replaced = Replace(order_notification_body_replaced,"%LASTNAME%",Last_name)
            order_notification_body_replaced = Replace(order_notification_body_replaced,"%FIRSTNAME%",First_name)
            order_notification_body_replaced = Replace(order_notification_body_replaced,"%LOGIN%",User_Id)
            order_notification_body_replaced = Replace(order_notification_body_replaced,"%PASSWORD%",Password)
            order_notification_body_replaced = Replace(order_notification_body_replaced,"%INVOICE_TEXT%",Send_Text_Invoice_By_Mail(loid)	)
	    end if
	    order_notification_subject_replaced=order_notification_subject
	    if InStr(order_notification_subject,"%") > -1 then
            order_notification_subject_replaced = Replace(order_notification_subject_replaced,"%ORDER%",loid)
            order_notification_subject_replaced = Replace(order_notification_subject_replaced,"%LASTNAME%",Last_name)
            order_notification_subject_replaced = Replace(order_notification_subject_replaced,"%FIRSTNAME%",First_name)
            order_notification_subject_replaced = Replace(order_notification_subject_replaced,"%LOGIN%",User_Id)
            order_notification_subject_replaced = Replace(order_notification_subject_replaced,"%PASSWORD%",Password)
            order_notification_subject_replaced = Replace(order_notification_subject_replaced,"%INVOICE_TEXT%",Send_Text_Invoice_By_Mail(loid)	)
	    end if
	    if order_notification_invoice then
	        fn_print_debug "in order_notification_invoice"
		    'FORMAT THE INVOICE TO BE SENT TO THE CUSTOMER
		    add_invoice = add_invoice_T1
		    Call Send_Invoice_By_Mail(loid,Shopper_Id,store_email,send_to,order_notification_subject_replaced,order_notification_body_replaced&chr(10)&add_invoice)
	    else
	        fn_print_debug "in else order_notification_invoice"
		    add_invoice = add_invoice_T2
		    Call Send_Mail_Html(Store_Email,send_to,order_notification_subject_replaced,order_notification_body_replaced&chr(10)&add_invoice)
	    end if
    end if
end sub


%>
