<%

rs_store.open "select score from store_maxmind WITH (NOLOCK) WHERE store_id="&store_id&" and oid="&oid&" and isValid=1"
if not rs_store.eof then
	score=rs_store("score")
else
	score=""
end if
rs_store.close

if score<>"" then
	'already checked fraud, dont check again
else

rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1
If not rs_store.eof then
	first_name =rs_store("first_name")
	last_name =rs_store("Last_name")
	zip =rs_store("Zip")
	country =rs_store("Country")
	phone = rs_store("Phone")
	email =rs_store("Email")
	address = rs_store("Address1") 
	address2 = rs_store("Address2") 
	city = rs_store("City")
	state = rs_store("State")
	ip =Request.ServerVariables("REMOTE_ADDR") 
	Invoice_Number = oid 
	Client_ID =cid 
	amount = formatnumber(GGrand_Total,2) 
End If
rs_store.close
sEmailSplit=split(email,"@")
if ubound(sEmailSplit)>0 then
	sDomain=sEmailSplit(1)
end if

	dim ccfs
	set ccfs = new CreditCardFraudDetection

	dim h
	set h = CreateObject("Scripting.Dictionary")

	'Enter your license key here
	'h.Add "license_key", "otPPDK5nNtAO"
	h.Add "license_key", maxmind_License
	

' Required fields
	h.Add "i", ip				' set the client ip address
	h.Add "city", city				' set the billing city
	h.Add "region", state					' set the billing state
	h.Add "postal", zip					' set the billing zip code
	h.Add "country", country					' set the billing country

	' Recommended fields
	h.Add "domain", sDomain			' Email domain
	h.Add "bin", left(CardNumber,6)
	'h.Add "binName", "MBNA America Bank"	' bank name
	'h.Add "binPhone", "800-421-2110"		' bank customer service phone number on back of credit card
	h.Add "custPhone", phone			' Area-code and local prefix of customer phone number
     h.Add "txnid", oid			' Area-code and local prefix of customer phone number

     h.Add "forwardedip", request.servervariables("HTTP_X_FORWARDED_FOR")
	h.Add "sessionid", Shopper_id
     h.Add "shipaddr", ShipAddress1
     h.Add "shipcity", ShipCity
     h.Add "shipregion", ShipState
     h.Add "shippostal", ShipZip
     h.Add "shipcountry", ShipCountry

     h.Add "emailmd5", fn_md5hash(email)
	h.Add "usernamemd5", fn_md5hash(user_id)
	h.Add "passwordmd5", fn_md5hash(password)




	ccfs.debug = 0
	if store_id=139 then
	ccfs.debug = 1

	end if
	ccfs.isSecure = 1
	ccfs.timeout = 5
	ccfs.input(h)
	ccfs.query()	 

	'Print out the result
	dim ret, outputkeys, numoutputkeys
	Set ret = ccfs.output()
	outputkeys = ret.Keys
	numoutputkeys = ret.Count

	sErr=""

	for i = 0 to numoutputkeys-1
		key = outputkeys(i)
		value = ret.Item(key)
		select case key
			case "distance"
				distance=value
			case "countryMatch"
				countryMatch=value
			case "countryCode"
				countryCode=value
			case "freeMail"
				freeMail=value
			case "anonymousProxy"
				anonymousProxy=value
			case "score"
				score=value
			case "binMatch"
				binMatch=value
			case "binCountry"
				binCountry=value
			case "err"
				serr=value				
			case "proxyScore"
				proxyScore=value
			case "spamScore"
				spamScore=value
			case "ip_region"
				ip_region=value
			case "ip_city"
				ip_city=value
			case "ip_latitude"
				ip_latitude=value
			case "ip_longitude"
				ip_longitude=value
			case "binName"
				binName=value
			case "ip_isp"
				ip_isp=value
			case "ip_org"
				ip_org=value
			case "binNameMatch"
				binNameMatch=value
			case "binPhoneMatch"
				binPhoneMatch=value
			case "binPhone"
				binPhone=value
			case "custPhoneInBillingLoc"
				custPhoneInBillingLoc=value
			case "highRiskCountry"
				highRiskCountry=value
			case "queriesRemaining"
				queriesRemaining=value
			case "cityPostalMatch"
				cityPostalMatch=value
			case "shipCityPostalMatch"
				shipCityPostalMatch=value
			case "maxmindID"
				maxmindID=value
			case "isTransProxy"
				isTransProxy=value
			case "carderEmail"
				carderEmail=value
			case "highRiskUsername"
				highRiskUsername=value
			case "highRiskPassword"
				highRiskPassword=value
			case "shipForward"
				shipForward=value
			case "riskScore"
				riskScore=value
			case "explanation"
				explanation=checkstringforQ(value)
		end select
	next
	if distance="" then
           distance=0
        end if
        if score="" then
	score=0
	end if
	if riskscore="" then
	riskscore=0
	end if
	if proxyscore="" then
	proxyscore=0
	end if
	if spamscore="" then
	spamscore=0
	end if
	if queriesremaining="" then
	queriesRemaining=0
	end if



        on error resume next
        sql_insert="insert into store_maxmind ([Store_Id], [oid], [anonymousProxy], [binCountry], [binMatch], [binName], [binNameMatch], [binPhone], [carderEmail], [cityPostalMatch], [countryMatch], [countryCode], [custPhoneInBillingLoc], [distance], [err], [explanation], [freeMail], [highRiskCountry], [highRiskPassword], [highRiskUsername], [ip_city], [ip_isp], [ip_latitude], [ip_longitude], [ip_org], [ip_region], [isTransProxy], [maxmindID], [proxyScore], [queriesRemaining], [riskscore], [score], [shipCityPostalMatch], [spamScore], [shipForward]) "&_
			"VALUES("&Store_Id&", "&oid&", '"&anonymousProxy&"','"&binCountry&"', '"&binMatch&"', '"&binName&"', '"&binNameMatch&"', '"&binPhone&"', '"&carderEmail&"', '"&cityPostalMatch&"', '"&countryMatch&"', '"&countryCode&"', '"&custPhoneInBillingLoc&"', "&distance&", '"&serr&"','"&explanation&"','"&freeMail&"', '"&highRiskCountry&"', '"&highRiskPassword&"', '"&highRiskUsername&"', '"&ip_city&"', '"&ip_isp&"', '"&ip_latitude&"', '"&ip_longitude&"', '"&ip_org&"', '"&ip_region&"', '"&isTransProxy&"', '"&maxmindID&"', "&proxyScore&", "&queriesRemaining&", "&riskScore&", "&score&", '"&shipCityPostalMatch&"', "&spamScore&", '"&shipForward&"')"

	   conn_store.execute sql_insert
end if

	if score<>"" then
            if cdbl(score)>=cdbl(maxmind_reject) then
    		fn_purchase_decline oid,"Your transaction has been declined by our fraud checking system."
            end if
	end if
%>
