
<%

' ================================================================ 
' RETURNS A STRING CONTAINING A CURRENT CURENCY SYMBOL AND FORMATED
' NUMERIC VALUE RECEIVED
Function Currency_Format_Function (Number)
Store_currency = Session("Store_currency")
	if IsNull(Number)  then 
		Number = 0
	end if	
	Currency_Format_Function = Store_currency&FormatNumber(Number,2)
End Function



%>
