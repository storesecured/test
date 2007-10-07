
<% 

'MATHEMATICAL VALIDATE A CC NUMBER
function IsCreditCard(ByRef asCardType, ByRef anCardNumber)

	Dim lsNumber
	Dim lsChar
	Dim lnTotal
	Dim lnDigit
	Dim lnPosition
	Dim lnSum

	IsCreditCard = False

	For lnPosition = 1 To Len(anCardNumber)
		lsChar = Mid(anCardNumber, lnPosition, 1)
    	if IsNumeric(lsChar) Then lsNumber = lsNumber & lsChar
	Next

	if Len(lsNumber) < 13 Then Exit function
	if Len(lsNumber) > 16 Then Exit function

	Select Case LCase(asCardType)
		Case "visa", "v"
	    	if Not Left(lsNumber, 1) = "4" Then Exit function
		Case "american express", "americanexpress", "american", "ax", "a"
    		if Not Left(lsNumber, 2) = "37" Then Exit function
		Case "mastercard", "master card", "master", "m"
			if Not Left(lsNumber, 1) = "5" Then Exit function
		Case "discover", "discovercard", "discover card", "d"
    		if Not Left(lsNumber, 1) = "6" Then Exit function
		Case Else
	End Select

	While Not Len(lsNumber) = 16
		lsNumber = "0" & lsNumber
	Wend

	For lnPosition = 1 To 16
    	lnDigit = Mid(lsNumber, lnPosition, 1)
	    lnMultiplier = 1 + (lnPosition Mod 2)
		lnSum = lnDigit * lnMultiplier
		if lnSum > 9 Then lnSum = lnSum - 9
		lnTotal = lnTotal + lnSum
	Next

	IsCreditCard = ((lnTotal Mod 10) = 0)

End function

%>
