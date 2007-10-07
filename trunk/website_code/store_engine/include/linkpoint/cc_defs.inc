<%
' Created from ccapi_error.h from API 3.8.3.17
Const Fail = 0
Const Succeed = 1

' Created from ccapi_client.h from API 3.8.3.17

'Request types possible for OrderField_Chargetype
Const Chargetype_Auth = 0
Const Chargetype_Sale = 0
Const Chargetype_Preauth = 1
Const Chargetype_Postauth = 2
Const Chargetype_Credit = 3
Const Chargetype_Voidsale = 9
Const Chargetype_Error = 0

'Result types possible for OrderField_Result
Const Result_Live = 0
Const Result_Good = 1
Const Result_Duplicate = 2
Const Result_Decline = 3

'ESD types for ItemField_Esdtype
Const Esdtype_None = 0
Const Esdtype_Softgood = 1
Const Esdtype_Key = 2

Const CVM_Not_Provided = "0"
Const CVM_Provided     = "1"
Const CVM_Illegible    = "2"
Const CVM_No_Imprint   = "9"
Const CVM_Not_Present  = " "


Const ECI_TRANSACTION    = "0"
Const MOTO_TRANSACTION   = "1"
Const RETAIL_TRANSACTION = "2"


Const TTYPE_UNSPECIFIED     = "0"
Const TTYPE_STANDALONE      = "1"
Const TTYPE_POS             = "2"
Const TTYPE_UNATTENDED      = "3"


' OrderField_t
Const OrderField_Oid = 0
Const OrderField_Userid = 1
Const OrderField_Bcompany = 2
Const OrderField_Bcountry = 3
Const OrderField_Bname = 4
Const OrderField_Baddr1 = 5
Const OrderField_Baddr2 = 6
Const OrderField_Bcity = 7
Const OrderField_Bstate = 8
Const OrderField_Bzip = 9
Const OrderField_Sname = 10
Const OrderField_Saddr1 = 11
Const OrderField_Saddr2 = 12
Const OrderField_Scity = 13
Const OrderField_Sstate = 14
Const OrderField_Szip = 15
Const OrderField_Scountry = 16
Const OrderField_Phone = 17
Const OrderField_Fax = 18
Const OrderField_Refer = 19
Const OrderField_Shiptype = 20
Const OrderField_Shipping = 21
Const OrderField_Tax = 22
Const OrderField_Subtotal = 23
Const OrderField_Vattax = 24
Const OrderField_Comments = 25
Const OrderField_PONumber = 26
Const OrderField_TaxExempt = 27
Const OrderField_MaxItems = 28
Const OrderField_Email = 29
Const OrderField_Cardnumber = 30
Const OrderField_Expmonth = 31
Const OrderField_Expyear = 32
Const OrderField_Chargetype = 33
Const OrderField_Chargetotal = 34
Const OrderField_Referencenumber = 35
Const OrderField_Result = 36
Const OrderField_Addrnum = 37
Const OrderField_Ip = 38
Const OrderField_MercType = 39
Const OrderField_SwitchStartDate = 40
Const OrderField_SwitchIssueNumber = 41
Const OrderField_CVMIndicator = 42
Const OrderField_CVMValue = 43
Const OrderField_MotoTransaction = 44
Const OrderField_RecurringTransaction = 45
Const OrderField_TDate = 46
Const OrderField_Track = 47
Const OrderField_TerminalType = 48

' Responses
Const OrderField_R_Time = 49
Const OrderField_R_Ref = 50
Const OrderField_R_Approved = 51
Const OrderField_R_Code = 52
Const OrderField_R_Ordernum = 53
Const OrderField_R_Error = 54
Const OrderField_R_FraudCode = 55
Const OrderField_R_AuthResponse = 56
Const OrderField_R_Message = 57
Const OrderField_R_TDate = 58

' ReqField_t
Const ReqField_Configfile = 0
Const ReqField_Keyfile = 1
Const ReqField_Appname = 2
Const ReqField_Host = 3
Const ReqField_Port = 4

' ItemField_t
Const ItemField_Itemid = 0
Const ItemField_Description = 1
Const ItemField_Price = 2
Const ItemField_Quantity = 3
Const ItemField_Softfile = 4
Const ItemField_Esdtype = 5
Const ItemField_Serial = 6
Const ItemField_MaxOptions = 7

' ShippingField_t
Const ShippingField_Country = 0
Const ShippingField_State = 1
Const ShippingField_Total = 2
Const ShippingField_Items = 3
Const ShippingField_Weight = 4
Const ShippingField_Carrier = 5
' Responses
Const ShippingField_R_Total = 6

' TaxField_t
Const TaxField_State = 0
Const TaxField_Zip = 1
Const TaxField_Total = 2
' Responses
Const TaxField_R_Total = 3

' OptionField_t
Const OptionField_Option = 0
Const OptionField_Choice = 1

' PbOrder_t
Const PbOrder_Submit = 0
Const PbOrder_Modify = 1
Const PbOrder_Cancel = 2
Const PbOrder_Process = 3

' PeriodicField_t
Const PeriodicField_Startdate = 0
Const PeriodicField_Periodicity = 1
Const PeriodicField_Installments = 2
Const PeriodicField_Threshold = 3
Const PeriodicField_Comments = 4

' CheckOrder_t
Const CheckOrder_Submit = 0
Const CheckOrder_Void = 1

' CheckField_t;
Const CheckField_AccountType = 0
Const CheckField_TransitRouting = 1
Const CheckField_MICR = 2
Const CheckField_CheckNumber = 3
Const CheckField_Comments = 4
const CheckField_SSN	        =5
const CheckField_DL		=6
const CheckField_DL_State       =7
const CheckField_BankName       =8
const CheckField_BankState	=9
const CheckField_Routing	=10
const CheckField_Account	=12
%>
