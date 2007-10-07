<!--#include virtual="common/connection.asp"-->

<%

file_name_path = "www/admin/include/sys_message.asp"

%>


<html>

<head>
	<title>Error !</title>
</head>

<script language="JavaScript">
	<!--
	function closeWin() {
		window.close();}
 
	if (window.Event) document.captureEvents(Event.ONCLICK);
		document.onclick = closeWin;
	// -->

</script>

<body>

<% 

Message_Add = fn_get_querystring("Message_Add")
Message_text = Not_defined_error 
select case fn_get_querystring("Message_id")
	case 5	Message_text = "Please do not try and change the address bar text, it contains important session information."
	case 6	Message_text = "Either you have attempted to order more items than we have in stock or we have reached our low quantity and cannot guarantee shipment, please try again with a lower quantity."
	case 7	Message_text = "Not a valid login, please try again."
	case 8	Message_text = "That userid already exists, please try again."
	case 9	Message_text = "Login information was not found in the database."
	case 10 Message_text = "The 2 passwords entered did not match."
	case 11 Message_text = "Coupon code entered is not valid or reached it maximum usage limit."
	case 13 Message_text = "We are sorry but your credit card information could not be verified with our financial instituation, please try again. <br>If you feel your credit card is valid, contact us immediatley to report the problem."
	case 14 Message_text = "Your search criteria must contain at least 3 characters."
	case 15 Message_text = "No store matching your request has been found."
	case 20 Message_text = "Sorry but we do not ship to the country specified."
	case 21 Message_text = "This field cannot be empty"
	case 22 Message_text = "This field must be an integer"
	case 24 Message_text = "Sorry but this store is no longer active."
	case 27 Message_text = "Quantity must be numeric."
	case 28 Message_text = "This field must be a date."
	case 29 Message_text = "Sorry no such cart exists."
	case 30 Message_text = "No valid records in database."
	case 31 Message_text = "First row cannot be larger than the last one."
	case 33 Message_text = "Email entered is not valid."
	case 34 Message_text = "Please select a valid region for the country specified, or if none match please select Other."
	case 35 Message_text = "Item selected is invalid."
	case 36 Message_text = "Quantity ordered must be at least "
	case 40 Message_text = "Unable to communicate with Verisign."
	case 41 Message_text = "Error reported by Verisign."
	case 42 Message_text = "You may only order whole items."
	case 43 Message_text = "Price entered was invalid."
	case 44 Message_text = "Price range is invalid."
	case 45 Message_text = "The credit card number entered is invalid."
	case 46 Message_text = "Your credit card has expired."
	case 47 Message_text = "Quantity cannot be blank."
	case 48 Message_text = "SKU and quantity cannot be blank."
	case 49 Message_text = "No item was found matching the SKU entered."
	case 50 Message_text = "There is more than one item with a SKU matching your entry."
	case 55 Message_text = "Your gift certificate has expired."
	case 56 Message_text = "Your gift certificate has already been fully used."
	case 57 Message_text = "This gift certificate can't be used with the following item: "
	case 64 Message_text = "The two passwords entered do not match."
	case 65 Message_text = "The affiliate code you entered is already in use, please try again."
	case 71 Message_text = "Please enter a valid address starting with http://"
	case 72 Message_text = "Please enter a valid email address."
	case 75 Message_text = "Invalid Order ID"
	case 81 Message_text = "You cancelled your payment or Paypal has reported an error.  Please try again."
	case 82 Message_text = "PsiGate Error. Error description:"
	case 85 Message_text = "Please input server address, user name, user password and access license for UPS server."
	case 86 Message_text = "Please input server address, user name and user password for USPS server."
	case 87 Message_text = "Please input server address for FedEx server."
	case 88 Message_text = "Please input server address for DHL server."
	case 89 Message_text = "To use real time shipping, please select UPS / USPS / DHL or FedEx."
	case 90 Message_text = "The database name cannot be empty."
	case 91 Message_text = "The connection string cannot be empty."
	case 92 Message_text = "Unable to connect to the database."
	case 93 Message_text = "Please select at least one value for each attribute."
	case 101 Message_text = ""
end select 
Message_text = Message_text&" "&Message_Add 

%> 

<table border="1" width="100%" cellspacing="0" height="100%">
	
	<tr> 
		<td width="100%" bgcolor="Red"><font face="Arial" size="2" color="White"><b><%= Error_T %>&nbsp;</b></font></td>
	</tr>
    
	<tr>
		<td width="100%" height="100%" valign="center" align="center"><font face="Arial" size="2">    
			<p><b><%= Message_text %></b></p></font></td>
    </tr>

    <tr bgcolor="Red">
	<th>
		<font face='arial' size='2' color='#ffffff'><%= Click_anywhere_in_this_window_to_close_it_T %></font>
	</th>
	</tr>
 
</table>

</body>
</html>
