<!--#include file="global_settings.asp"-->



<%
		
' coming from ups_print_labels.asp after sending the accept request
if Request.Form("RefNo") <> "" then
    Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")
															
'if nodeResponseCode.text <> 0 then	 ' if confirm request is successful (Code=0)
				
		RefNo=Request.Form("RefNo")

		sql_real_time = "select * from Store_Real_Time_Settings where store_id="&store_id
		rs_Store.open sql_real_time, conn_store, 1, 1
			UPS_User = rs_store("UPS_User")
			UPS_Password = rs_store("UPS_Password")
			UPS_AccessLicense = rs_store("UPS_AccessLicense")
			UPS_Account_Number = rs_store("UPS_Account_Number")
		rs_Store.Close

		'getting the access info
		Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.load(Server.MapPath( "xml/access.xml")) 
		
		set nodeUPS_User = xmlDoc.documentElement.selectsinglenode("UserId")
		nodeUPS_User.text=UPS_User
    'nodeUPS_User.text="melanieeasystore"

		set nodeUPS_Password = xmlDoc.documentElement.selectsinglenode("Password")
		nodeUPS_Password.text=UPS_Password
		'nodeUPS_Password.text="ankle237"

		set nodeUPS_AccessLicense = xmlDoc.documentElement.selectsinglenode("AccessLicenseNumber")
		nodeUPS_AccessLicense.text=UPS_AccessLicense
		'nodeUPS_AccessLicense.text="FBBA98A10FDA0F48"

		strAccess=xmlDoc.xml	

		Set oConfirm= Server.CreateObject("MSXML2.DOMDocument")
		oConfirm.load(Upload_Folder&"ups_confirm_op_"&RefNo&".xml")

		'extract the ShipDigest sent in the confirm response
		set nodeDigest=oConfirm.documentElement.selectsinglenode("ShipmentDigest")
		strDigest=nodeDigest.text

		'create the accept request
		Set xmlDocReq= Server.CreateObject("MSXML2.DOMDocument")
		xmlDocReq.load(Server.MapPath( "xml/ups_accept.xml")) 
		set nodeDigest=xmlDocReq.documentElement.selectsinglenode("ShipmentDigest")
		nodeDigest.text=strDigest
		
		'save the accept request after making changes
		xmlDocReq.Save(Upload_Folder&"ups_accept_"&RefNo&".xml") 

		Set xmlDocReq= Server.CreateObject("MSXML2.DOMDocument")
		xmlDocReq.load(Upload_Folder&"ups_accept_"&RefNo&".xml") 
		strReq=xmlDocReq.xml

		'send the accept request			
		Set xmlHTTP = Server.CreateObject("Microsoft.XMLHTTP")
		xmlHTTP.open "POST","https://www.ups.com/ups.app/xml/ShipAccept", False
		xmlHTTP.send(strAccess+strReq)

		'save the accept response
		Set oConfirm= Server.CreateObject("MSXML2.DOMDocument")
		oConfirm.loadXML(xmlHTTP.ResponseText)
		oConfirm.save(Upload_Folder&"ups_accept_op_"&RefNo&".xml")
		
		fn_redirect "labels_list.asp"
end if

%>
