<!--#include file="sys_countries.asp"-->

<%

function create_country_list (sFieldName,sValue,sSize,sAddList)
	if sSize>1 then
		sSize=sSize&" multiple"
	end if
	sText="<select name='"&sFieldName&"' size="&sSize&">"
	sList=sAddList&sSysCountries
	FOR each sCountryRow in split(sList,":")
		sSelected = ""
		if (sValue<>"") then
			if Is_In_Collection(sValue,sCountryRow,",") then					
			  sSelected = "selected"
			end if
		end if
			
		sText = sText&"<option "&sSelected&" value='"&sCountryRow&"'>"&sCountryRow
	Next
	sText=sText & "</select>"
	create_country_list=sText
end function

%>