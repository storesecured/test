
<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%
	
	
if Request.Querystring("Id") <> ""  then

	  RefNo=Request.Querystring("Id")

	  Set FileObject = CreateObject("Scripting.FileSystemObject")
	  Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")
	  
		if  FileObject.FileExists(Upload_Folder&"/ups_accept_op_"&RefNo&".xml") then
		
				Set xmlDoc= Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.load(Upload_Folder&"/ups_accept_op_"&RefNo&".xml")
				

				'exracting the shipment id from the stored accept response
				set nodeShipmentIdNumber = xmlDoc.documentElement.selectsinglenode("ShipmentResults/ShipmentIdentificationNumber")
				strShipmentIdNumber=nodeShipmentIdNumber.text			
			

				'fetch UPS details
				sql_real_time = "select * from Store_Real_Time_Settings where store_id="&store_id
				rs_Store.open sql_real_time, conn_store, 1, 1
					UPS_User = rs_store("UPS_User")
					UPS_Password = rs_store("UPS_Password")
					UPS_AccessLicense = rs_store("UPS_AccessLicense")
					UPS_Account_Number = rs_store("UPS_Account_Number")
					UPS_Pack=rs_store("UPS_Pack")
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


				'sending the void request			
				Set xmlDocReq= Server.CreateObject("MSXML2.DOMDocument")
				xmlDocReq.load(Server.MapPath( "xml/ups_void.xml")) 

				set nodeShipmentIdentificationNumber=xmlDocReq.documentElement.selectsinglenode("ShipmentIdentificationNumber")
				nodeShipmentIdentificationNumber.text=strShipmentIdNumber

				strReq=xmlDocReq.xml

				'  combining access + void xml and sending the void request
				Set xmlHTTP = Server.CreateObject("Microsoft.XMLHTTP")
				xmlHTTP.open "POST","https://www.ups.com/ups.app/xml/Void", False
				xmlHTTP.send(strAccess+strReq)
		

				'saving the void response to an xml file	
				Set oVoid= Server.CreateObject("MSXML2.DOMDocument")
				oVoid.loadXML(xmlHTTP.ResponseText)
				oVoid.save(Upload_Folder&"/ups_void_op_"&RefNo&".xml")


				set nodeResponseCode= oVoid.documentElement.selectsinglenode("Response/ResponseStatusCode")

				if nodeResponseCode.text <> 0 then	 ' if void request is successful (Code=1)
					set nodeResponseStatusDescription= oVoid.documentElement.selectsinglenode("Response/ResponseStatusDescription")
						fn_redirect "labels_list.asp"
				else
						set nodeError=oVoid.documentElement.selectsinglenode("Response/Error/ErrorDescription")
						 fn_error nodeError.text
				end if

		else
			fn_error "Label does not exist"
		end if

end if
%>
