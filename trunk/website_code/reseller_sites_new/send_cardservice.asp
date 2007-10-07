
<%
if request.form("recipient") <> "" then
    name1=Request.Form("name1")
    name2=Request.Form("name2")
    name3=Request.Form("name3")
    address4=Request.Form("address4")
    telephone5=Request.Form("telephone5")
    city6=Request.Form("city6")
    bstate=Request.Form("bstate")
    zip8=Request.Form("zip8")
    email9=Request.Form("email9")
    recipient=Request.Form("recipient")

    message = name1 & " " & name2 & vbcrlf &_
        name3 & vbcrlf & address4 & vbcrlf &_
        telephone5 & vbcrlf & city6 & "," & _
        bState & " " & zip8 & vbcrlf & email9

    Set Mail = Server.CreateObject("Persits.MailSender")
    Mail.From = email9
    Mail.AddAddress sReport_email
    Mail.AddAddress recipient
    Mail.Subject = "LinkPoint Lead From CD63"
    Mail.Body = message
    On Error Resume Next
    Mail.Send


elseif request.form("recipientecho") <> "" then
    legalname=vbcrlf&"Legal Name " & Request.Form("legalname")
    dbaname=vbcrlf&"DBA Name " & Request.Form("dbaname")
    firstname=vbcrlf&"First Name "&Request.Form("firstname")
    lastname=vbcrlf&"Last Name " &Request.Form("lastname")
    address4=vbcrlf&"Business Address "&Request.Form("address4")
    telephone5=vbcrlf&"Telephone "&Request.Form("telephone5")
    city6=vbcrlf&"City " &Request.Form("city6")
    bstate=vbcrlf&"State " &Request.Form("bstate")
    zip8=vbcrlf&"Zip "&Request.Form("zip8")
    email9=Request.Form("email9")
    recipientecho=Request.Form("recipientecho")
    type_of_business=vbcrlf&"Type of Business "&Request.Form("type_of_business")
    products_sold=vbcrlf&"Products Sold "&Request.Form("products_sold")
    echeck=vbcrlf&"Use eCheck "&Request.Form("echeck")

    message = legalname & " " & dbaname & vbcrlf &_
        firstname & lastname & vbcrlf & address4 & vbcrlf &_
        telephone5 & vbcrlf & city6 & "," & _
        bState & " " & zip8 & vbcrlf & email9  & vbcrlf&_
        type_of_business & products_sold & echeck

    Set Mail = Server.CreateObject("Persits.MailSender")
    Mail.From = email9
    Mail.AddAddress recipientecho
    Mail.Subject = "Echo Lead"
    Mail.Body = message
    On Error Resume Next
    Mail.Send

end if
%>
You will be contacted shortly.
