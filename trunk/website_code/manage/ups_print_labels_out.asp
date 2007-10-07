<!--#include file="Global_Settings.asp"-->
<%
	
	
if Request.Querystring("Id") <> ""  then
    Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")
    RefNo=Request.Querystring("Id")

	  Set FileObject = CreateObject("Scripting.FileSystemObject")
		if  FileObject.FileExists(Upload_Folder&"ups_accept_op_"&RefNo&".xml") then
		
				Set xmlDoc= Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.load(Upload_Folder&"ups_accept_op_"&RefNo&".xml")
				
				'extract the base 64 encoded graphic of the label
				set nodeGraphic = xmlDoc.documentElement.selectsinglenode("ShipmentResults/PackageResults/LabelImage/GraphicImage")
				strText=nodeGraphic.text

				DeliveryConfirmationImage= strText
				strDeliveryConfirmationLabel=strText

				'decode base64 encoding
				  Dim dcOutput()
				  sBase ="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
				  StreamLength = len(strDeliveryConfirmationLabel)
				  '---------------------------------------------
				  'Strip cr/lf
				  dcCopylen = 0
				  dcCopy = ""
				  For j = 1 To StreamLength
					If ((Mid(strDeliveryConfirmationLabel, j, 1) <> vbCr) _
					  And (Mid(strDeliveryConfirmationLabel, j, 1) <> vbLf)) Then
					  dcCopy = dcCopy & Mid(strDeliveryConfirmationLabel, j, 1)
					  dcCopylen = dcCopylen + 1
					End If
				  Next
				  '----------------------------------------------
				  '----------------------------------------------
				  'Decode bulk of string
				  dcOutputlen = 0
				  ReDim dcOutput((dcCopylen * 3) / 4)
				  For j = 1 To dcCopylen - 4 Step 4
					'map "A"-"/" to 0-63
					a = InStr(1, sBase, Mid(dcCopy, j, 1), vbBinaryCompare) - 1
					b = InStr(1, sBase, Mid(dcCopy, j + 1, 1), vbBinaryCompare) -1
					c = InStr(1, sBase, Mid(dcCopy, j + 2, 1), vbBinaryCompare) -1
					d = InStr(1, sBase, Mid(dcCopy, j + 3, 1), vbBinaryCompare) -1
					'decode 0-63 to 0-255
					dcOutput(dcOutputlen) = (a * 4) Or ((b And 48) / 16)
					dcOutput(dcOutputlen + 1) = ((b And 15) * 16) Or ((c And 60) / 4)
					dcOutput(dcOutputlen + 2) = ((c And 3) * 64) Or (d And 63)
					dcOutputlen = dcOutputlen + 3
				  Next
				  '-----------------------------------------------
				  '-----------------------------------------------
				  'Decode last 1-3 characters
				  a = InStr(1, sBase, Mid(dcCopy, j, 1), vbBinaryCompare) - 1
				  b = InStr(1, sBase, Mid(dcCopy, j + 1, 1), vbBinaryCompare) - 1
				  dcOutput(dcOutputlen) = (a * 4) Or ((b And 48) / 16)
				  If j + 2 <= dcCopylen Then
					c = InStr(1, sBase, Mid(dcCopy, j + 2, 1), vbBinaryCompare) - 1
					dcOutput(dcOutputlen + 1) = ((b And 15) * 16) Or ((c And 60) / 4)
					dcOutputlen = dcOutputlen + 1
				  End If
				  If j + 3 <= dcCopylen Then
					d = InStr(1, sBase, Mid(dcCopy, j + 3, 1), vbBinaryCompare) - 1
					dcOutput(dcOutputlen + 2) = ((c And 3) * 64) Or (d And 63)
					dcOutputlen = dcOutputlen + 1
				  End If
				  
				Response.ContentType = "image/GIF"  
				for j = 0 to dcOutputlen - 1
					Response.BinaryWrite chrB(dcOutput(j))
				next

		
		else
			Response.write"<BR> <font face=verdana size=2 color=red> Label does not exist </font>"
		end if

end if
%>
