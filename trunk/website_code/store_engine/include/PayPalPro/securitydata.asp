
<%

'----------------------------------------------------------------------------------
' PayPal Constants File
' ====================================
' Authentication Credentials for making the call to the server
'----------------------------------------------------------------------------------

	set rs_Store =  server.createobject("ADODB.Recordset")
	if request.form("store_id")<>"" then
	 store_id=  request.form("store_id")
        end if
sql_real_time = "exec wsp_real_time_property "&Store_Id&",36;"
rs_Store.open sql_real_time,conn_store,1,1
rs_Store.MoveFirst
Do While Not Rs_Store.EOF 
    select case Rs_store("Property")
    case "PayPal_Pro_Api_Username"
	    PayPal_Pro_API_username = decrypt(Rs_store("Value"))
    case "PayPal_Pro_Password"
	    PayPal_pro_Password = decrypt(Rs_store("Value"))
    case "PayPal_Pro_Currency"
        PayPal_Pro_Currency = decrypt(Rs_store("Value"))
    case "PayPal_Pro_Signature"
        PayPal_Pro_Signature = decrypt(Rs_store("Value"))
     end select
Rs_store.MoveNext
Loop
Rs_store.Close
set rs_Store = nothing

 API_USERNAME	= PayPal_Pro_API_username
 API_PASSWORD	= PayPal_pro_Password
 API_SIGNATURE	= PayPal_Pro_Signature
	
	'if Store_Id=33019 or Store_Id= 32084 then
         'API_ENDPOINT	= "https://api-3t.sandbox.paypal.com/nvp"
         'API_USERNAME	= "bassel_gado_api1.yahoo.com"
	 'API_PASSWORD	= "EBJSQG3FDTW8KEAH"
	 'API_SIGNATURE	= "AGrpXtwFJ7TMeqHT6t2t-5rWd7u1Afcz3IcSFAbsB4nALeUtir-cH6DV"

        'else
 	  Const API_ENDPOINT    = "https://api-3t.paypal.com/nvp"
	'end if
	Const API_VERSION	= "2.3"
	 PAYPAL_EC_URL	= "https://www.sandbox.paypal.com/webscr"

	Const HTTPREQUEST_PROXYSETTING_SERVER = ""
	Const HTTPREQUEST_PROXYSETTING_PORT = ""
	Const USE_PROXY = False

%>
