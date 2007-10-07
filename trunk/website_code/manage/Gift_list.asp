<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
if request.querystring("Id") ="" then
   response.redirect "error.asp?Message_Id=1"
end if

'SEND GIFT CERTIFICATE CODE BY MAIL TO CUSTOMER AND EMAIL ADDRESS SPECIFIED
'BY CUSTOMER

on error goto 0
if request.queryString("Mail")<>"" then
	sMail=request.queryString("Mail")
  	if not isNumeric(sMail) then
  		Response.Redirect "admin_error.asp?message_id=1"
  	end if
  	call sub_mailCertificate(sMail)
  	response.redirect "gift_list.asp?Id="&request.querystring("Id")
end if

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Gift_Certificates_Dets"
myStructure("TableWhere") = "gift_id="&request.querystring("Id")
myStructure("ColumnList") = "gift_det_id,gift_code,order_id,initial_amount,current_amount"
myStructure("HeaderList") = "gift_code,order_id,initial_amount,current_amount,mail"
myStructure("DefaultSort") = "gift_code"
myStructure("PrimaryKey") = "gift_det_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 0
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = "gift_manager.asp"
myStructure("BackToName") = "Gift Certificate"
myStructure("Menu") = "marketing"
myStructure("FileName") = "gift_list.asp?Id="&request.querystring("Id")
myStructure("FormAction") = "gift_list.asp"
myStructure("Title") = "Gift Certificate Purchases"
myStructure("FullTitle") = "Marketing > <a href=gift_manager.asp class=white>Gift Certificates</a> > Purchases"
myStructure("CommonName") = "Gift Certificate Purchase"
myStructure("NewRecord") = "gift_add.asp"
myStructure("Heading:gift_det_id") = "PK"
myStructure("Heading:order_id") = "Order"
myStructure("Format:order_id") = "STRING"
myStructure("Link:order_id") = "order_details.asp?Id=THISFIELD"
myStructure("Heading:gift_code") = "Gift Code"
myStructure("Format:gift_code") = "STRING"
myStructure("Heading:initial_amount") = "Initial"
myStructure("Format:initial_amount") = "CURR"
myStructure("Heading:current_amount") = "Current"
myStructure("Format:current_amount") = "CURR"
myStructure("Heading:mail") = "Mail"
myStructure("Format:mail") = "TEXT"
myStructure("Link:mail") = "THISURLMail=PK"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
	if Request.QueryString("Delete_Id") <> "" then

	end if

createFoot thisRedirect, 0


sub sub_mailCertificate(gift_id)
	sql_select = "select gift_email_subject, gift_buyer_email_body from store_emails WITH (NOLOCK) where store_id="&Store_id
	set rs_gift=server.createobject("adodb.recordset")
	rs_gift.open sql_select, conn_store, 1, 1
     Gift_email_Subject=rs_gift("Gift_email_Subject")
     gift_buyer_email_body=rs_gift("gift_buyer_email_body")
     rs_gift.close
     
     sql_select = "select * from store_gift_certificates_dets WITH (NOLOCK) where store_id="&store_id&" and gift_det_id="&gift_id
     rs_gift.open sql_select, conn_store, 1, 1
     	gift_code=rs_gift("gift_code")
          gift_amount=rs_gift("initial_amount")
		mail_to=rs_gift("mail_to")
		mail_text=rs_gift("mail_text")
     rs_gift.close
     set rs_gift=nothing
	send_cont = Gift_buyer_email_body
	if instr(send_cont,"%GIFT_CODE%")>0 then
     	send_cont=replace(send_cont,"%GIFT_CODE%",gift_code)
     else
          send_cont = send_cont&vbcrlf&"Your gift certificate code is: "&gift_code
     end if
     send_cont=replace(send_cont,"%AMOUNT%",Currency_Format_Function(gift_amount))

     send_cont = send_cont&vbcrlf&mail_text

     call Send_Mail_Html(store_email, mail_to, Gift_email_Subject, send_cont)

	sql_update="update store_gift_certificates_dets set mailed=1 where store_id="&store_id&" and gift_det_id="&gift_id
	conn_store.execute sql_update

end sub
%>

