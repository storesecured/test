
<%

dim g_KeyLocation
dim g_KeyLen 
Dim g_Key 

g_KeyLocation = fn_get_code_folder("Certificates")&"CCkey.key"

g_KeyLen = 512

Function EnCrypt(strCryptThis)
	if isNull(strCryptThis) then
		EnCrypt = strCryptThis
	Else
	  Dim strChar, iKeyChar, iStringChar, i
	  g_Key = mid(ReadKeyFromFile(g_KeyLocation),1,Len(strCryptThis))

	  for i = 1 to Len(strCryptThis)
		  iKeyChar = Asc(mid(g_Key,i,1))
		  iStringChar = Asc(mid(strCryptThis,i,1))
		  iCryptChar = iKeyChar Xor iStringChar
		  strEncrypted =  strEncrypted & createCrStr(iCryptChar)
	  next
	  EnCrypt = strEncrypted
	End if
End Function

Function DeCrypt(strEncrypted)
	if isNull(strEncrypted) then
	    DeCrypt = strEncrypted
	else
	  Dim strChar, iKeyChar, iStringChar, i
	  g_Key = mid(ReadKeyFromFile(g_KeyLocation),1,Len(strEncrypted))
	  for i = 1 to Len(strEncrypted) step 3
		  iKeyChar = Asc(mid(g_Key,(i-1)/3+1,1))
		  iStringChar = mid(strEncrypted,i,3)
		  iDeCryptChar = iKeyChar Xor iStringChar
		  strDecrypted =  strDecrypted & Chr(iDeCryptChar)
	  next
	  DeCrypt = strDecrypted
	end if
End Function

Function ReadKeyFromFile(strFileName)
	Dim keyFile, fso, f
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set f = fso.GetFile(strFileName)
	set ts = f.OpenAsTextStream(1, -2)
	Do While not ts.AtEndOfStream
	keyFile = keyFile & ts.ReadLine
	Loop

	set fso = Nothing
	ReadKeyFromFile =  keyFile
End Function

Sub WriteKeyToFile(MyKeyString,strFileName)
	Dim keyFile, fso
	set fso = Server.CreateObject("scripting.FileSystemObject") 
	set keyFile = fso.CreateTextFile(strFileName, true) 
	keyFile.WriteLine(MyKeyString)
	keyFile.Close
	set fso = Nothing
End Sub

Function KeyGeN(iKeyLength)
	Dim k, iCount, strMyKey
	lowerbound = 35
	upperbound = 96
	Randomize
	for i = 1 to iKeyLength
		s = 255
		k = Int(((upperbound - lowerbound) + 1) * Rnd + lowerbound)
		strMyKey =  strMyKey & Chr(k) & ""
	next
	KeyGeN = strMyKey
End Function

function createCrStr(nval)
	dim retCrVal
	retCrVal = cstr(abs(nval))
	do while len(retCrVal)<3
		retCrVal = "0"&retCrVal
	loop
	createCrStr = retCrVal
end function

%>
