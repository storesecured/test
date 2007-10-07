<!--#include file="include/header.asp"-->
<%

function fn_get_error (sQString)
	sSanitize=request.querystring(sQString)
	sSanitize=fn_replace(sSanitize,"<script","")
     sSanitize=fn_replace(sSanitize,"</script>","")
     sSanitize=fn_replace(sSanitize,"<iframe","")
     sSanitize=fn_replace(sSanitize,"src=","")

     fn_get_error=sSanitize
end function

Message_Add = fn_get_error("Message_Add")
Message_text = Not_defined_error

iMessage_Id = fn_get_querystring("Message_id")
if isNumeric(iMessage_Id) then
select case iMessage_Id
	case 1	Message_text = "Invalid data entry, request cancelled."
	case 2	Message_text = "Data was not submitted from our site, request cancelled."
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
else
    if iMessage_Id <> "" then
       Message_text = "Invalid Error Id"
    end if
end if

if fn_get_querystring("Item_Name")<>"" then
   Message_Add="<BR>Item="&fn_get_querystring("Item_Name")
end if
if instr(Message_Add,"|")>0 then
	Message_text=Message_text & "<ul>"
	for each one_error_array in split(Message_Add,"|")
		if len(one_error_array) > 1 then
			Message_text=Message_text&"<li>"&one_error_array&"</li>"
		End If
	Next
	Message_text=Message_text & "</ul>"
else
	Message_text = Message_text&" "&Message_Add
end if
%> 

<table border="0" width="100%" cellspacing="0" height=1>
	

	 
	<tr>
		<td valign=top width="100%" align="center"><font face="Arial" size="2">
			<p><b>Error<BR><BR><%= Message_text %></b><BR><BR>Use your browser's back button to return to the previous page.</p></font></td>
	 </tr>


 
</table>
<!--#include file="include/footer.asp"-->
