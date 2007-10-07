<%
Class clsWhois

	Public URL
	Public CropAfter
	Public CropBefore

	Private Sub Class_Initialize()
		URL = "http://netsol.com/cgi-bin/whois/whois?STRING="
		CropAfter = "<!--ingoes pre-->"
		CropBefore = "<!-- out goes pre-->"
	End Sub
	
	Function LookUp(pStrDomain)

		Dim lStrURL
		Dim lObjRegExp
		Dim lStrLookUp
		Dim strOutcome

		If pStrDomain = "" Then Exit Function

		lStrLookUp = GetURL(URL &Server.URLEncode(pStrDomain) &"&SearchType=do")

		Dim lLngStart
		Dim lLngEnd
		
		If Not(CropAfter = "" Or CropBefore = "") Then
			lLngStart = InStr(1, lStrLookUp, CropAfter, vbTextCompare)
			If lLngStart = 0 Then
				lStrLookUp = ""
			Else
				lLngStart = lLngStart + Len(CropAfter) + 1
				lLngEnd = InStr(lLngStart, lStrLookUp, CropBefore, vbTextCompare)
				If lLngEnd = 0 Then
					lStrLookUp = ""
				Else
					lLngEnd = lLngEnd - lLngStart
					lStrLookUp = Mid(lStrLookUp, lLngStart, lLngEnd)
				End If
			End If
		End If

		' remove HTML tags
		Set lObjRegExp = New RegExp
		lObjRegExp.Global = True
		lObjRegExp.Pattern = "<[^>]*>"
		lStrLookUp = lObjRegExp.Replace(lStrLookUp, "")
		Set lObjRegExp = Nothing
		
		lLngStart = InStr(1, lStrLookUp, "This name is available for registration", vbTextCompare)
		if lLngStart <> 0 then
			strOutcome = "1"
		else
			lLngStart = InStr(1, lStrLookUp, "WHOIS Record for", vbTextCompare)
			if lLngStart <> 0 then
				strOutcome = "0"
			else
				strOutcome = lStrLookUp
			end if
		end if

		LookUp = strOutcome

	End Function

	Public Function GetURL(ByRef pStrURL)

		Dim lObjSpider
		Dim strText
		If pStrURL = "" Then Exit Function


		' Different variations of XML objects
		set lObjSpider = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")

		With lObjSpider
			.Open "GET", pStrURL, False, "", ""
			.Send
			GetURL = .ResponseText
		End With
		Set lObjSpider = Nothing

	End Function
	
End Class
%>
