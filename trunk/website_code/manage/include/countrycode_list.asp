<!--#include file="sys_countries_code.asp"-->

<%

function create_countrycode_list (sFieldName,sValue,sSize,sAddList)
	on error goto 0
	if sSize>1 then
		sSize=sSize&" multiple"
	end if
	sText="<select name='"&sFieldName&"' size="&sSize&">"
	sList=sAddList&sCountriesCode
	FOR each sCountryRow in split(sList,"|")
		if sCountryRow<>"" then
			sCountryArray=split(sCountryRow,";")
			sCountry_Id=sCountryArray(0)
			sCountry_Name=sCountryArray(1)
			sSelected = ""
			if (sValue<>"") then
				if Is_In_Collection(sValue,sCountry_Id,",") then					
				  sSelected = "selected"
				end if
			end if
			
			sText = sText&"<option "&sSelected&" value='"&sCountry_Id&"'>"&sCountry_Name
		end if
	Next
	sText=sText & "</select>"
	create_countrycode_list=sText
end function

%>