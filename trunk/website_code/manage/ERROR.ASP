
<html>

<head>
   <title>Error</title>
</head>


<body>

<%

Message_Add = request("Message_Add")
Message_text = Not_defined_error
Message_Id= request("Message_Id")
if isNumeric(Message_Id) then
select case Message_Id
   case 1 Message_text = "Invalid data entry, request cancelled."
   case 2   Message_text = "Data was not submitted from our site, request cancelled."
   case 7   Message_text = "You have entered an invalid Site #, Username or Password, please try again."
   case 8   Message_text = "That login already exists, please try again."
   case 9   Message_text = "Login information was not found in the database."
   case 10 Message_text = "The 2 passwords entered did not match."
   case 12 Message_text = "You must accept at least one method of payment."
   case 13 Message_text = "We are sorry but your credit card information could not be verified with our financial instituation, please try again. <br>If you feel your credit card is valid, contact us immediatley to report the problem."
   case 15 Message_text = "No store matching your request has been found."
   case 21 Message_text = "This field cannot be empty"
   case 22 Message_text = "This field must be an integer"
   case 23 Message_text = "You must select a theme"
   case 24 Message_text = "Sorry but this store is no longer active."
   case 25 Message_text = "This coupon id already exists."
   case 26 Message_text = "There is no activity for the selected period"
   case 27 Message_text = "Quantity must be numeric."
   case 28 Message_text = "This field must be a date."
   case 29 Message_text = "Sorry no such cart exists."
   case 30 Message_text = "No valid records in database."
   case 31 Message_text = "First row cannot be larger than the last one."
   case 32 Message_text = "The specified page already exists, please try again."
   case 33 Message_text = "Email entered is not valid."
   case 34 Message_text = "Please select a valid region for the country specified, or if none match please select Other."
   case 35 Message_text = "Item selected is invalid."
   case 42 Message_text = "You may only order whole items."
   case 43 Message_text = "Price entered was invalid."
   case 44 Message_text = "Price range is invalid."
   case 47 Message_text = "Quantity cannot be blank."
   case 48 Message_text = "SKU and quantity cannot be blank."
   case 49 Message_text = "No item was found matching the SKU entered."
   case 50 Message_text = "There is more than one item with a SKU matching your entry."
   case 51 Message_text = "The Gift Certificate name selected is already in use, please try again."
   case 52 Message_text = "Gift Certificate amount must be greater than 0."
   case 53 Message_text = "Gift Certificate validity must be greater than 0."
   case 54 Message_text = "The item you have selected is already a gift certificate please use a different item."
   case 58 Message_text = "The data base already contains order records with a value above the selected value, the minimum value you can specify for Starting Order ID # is "
   case 59 Message_text = "The data base already contains item records with a value above the selected value, the minimum value you can specify for Starting Item ID # is "
   case 60 Message_text = "The data base already contains transaction records with a value above the selected value, the minimum value you can specify for Starting Transaction ID # is "
   case 61 Message_text = "The data base already contains customer records with a value above the selected value, the minimum value you can specify for Starting Customer ID # is "
   case 64 Message_text = "The two passwords entered do not match."
   case 65 Message_text = "The affiliate code you entered is already in use, please try again."
   case 71 Message_text = "Please enter a valid address starting with http://"
   case 72 Message_text = "Please enter a valid email address."
   case 75 Message_text = "Invalid Order ID"
   case 85 Message_text = "Please input server address, user name, user password and access license for UPS server."
   case 86 Message_text = "Please input server address, user name and user password for USPS server."
   case 87 Message_text = "Please input server address for FedEx server."
   case 88 Message_text = "Please input server address for DHL server."
   case 89 Message_text = "To use real time shipping, please select UPS / USPS / DHL or FedEx."
   case 90 Message_text = "The database name cannot be empty."
   case 91 Message_text = "The connection string cannot be empty."
   case 92 Message_text = "Unable to connect to the database."
   case 93 Message_text = "Please select at least one value for each attribute."
   case 100 Message_text = Session("Long_Error_Text")
   case 101 Message_text = ""
   case 102 Message_text = "You cannot save a template without a center content keyword.<BR><BR>Please add %OBJ_CENTER_CONTENT_OBJ% keyword tag where you would like the center content to be."
   case 103 Message_text = "Before selecting move please ensure that you have selected at least one item and a department."
   case 104 Message_text = "The site name choosen already exists, please try again."
   case 105 Message_text = "Please select the Item first to which the Price Group needs to be created. <br> Please Click <a href='edit_items.asp'>here</a> to select the Item "
end select
end if
Message_text = Message_text&" "&Message_Add

%>

<table border="1" width="380" cellspacing="0">
   <tr>
      <td width="100%" bgcolor="Red"><font face="Arial" size="2" color="White"><b>Error </b></font></td>
   </tr>


   <tr>
		 <td width="100%" height="100%" valign="center" align="center"><font face="Arial" size="2">

         <ul>
         <LI><%= Message_text %></LI></UL>

         <% if message_id = "7" then%>
         <a href=http://manage.storesecured.com>Return to Login Screen</a>
         <br><br><font face="Arial" size="2">
         <a href=forgot_password.asp class=link>Forgotten your Site #, Username or Password</a>
         </font>
         <% else %>

         Use your browsers back button to return.</font>
         <% end if%>
         </td>
    </tr>
    

</table>

</body>
</html>


