<%

function fn_make_location_list()
	'build list up from db this one time so we dont need to later
	sql_select = "SELECT Ship_Location_Id,Location_Name FROM Store_Ship_Location WITH (NOLOCK) where store_id="&Store_Id&" ORDER BY Location_Name"

   set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	sList=""
	if noRecords=0 then
		FOR myrowcounter= 0 TO myfields("rowcount")
			sList = sList & (mydata(myfields("ship_location_id"),myrowcounter)&"|{}|"&_
			mydata(myfields("location_name"),myrowcounter)&"|{new}|")
		next
	else
		sList="Empty"
	end if
	fn_make_location_list=sList
end function

function create_location_list (sFieldName,sValue,sSize)
    sLocationList = fn_make_location_list
	if sLocationList="Empty" then
		sText="<input type=hidden name='"&sFieldName&"'	value=0>Default"
	else
		if sSize>1 then
			sSize=sSize&" multiple"
		end if
		sText="<select name='"&sFieldName&"' size="&sSize&">"
		if (sValue<>"") then
			if sValue="0" then
				sSelected="selected"
			else
				sSelected =""
			end if
		end if
		sText=sText&"<Option value='' "&sSelected&" >Please Select</option>"
      sText =sText&"<Option value='0' "&sSelected&">Default"

		for each sLocationRecord in split(sLocationList,"|{new}|")
			if sLocationRecord<>"" then
				sLocationArray=split(sLocationRecord,"|{}|")
				sLocation_Id=sLocationArray(0)
				sLocation_Name=sLocationArray(1)
				sSelected=""
				if (sValue<>"") then
					if cdbl(sValue)=cdbl(sLocation_Id) then			
						sSelected = "selected"
					End If
				end if
				sText=sText&"<Option value='"&sLocation_Id&"' "&sSelected&" >"&sLocation_Name&"</option>"
				
			end if
		next
		
		sText=sText&"</select><input type='hidden' name='"&sFieldName&"_C' value='Re|String|||||Location'>"
	end if
	create_location_list=sText
end function

%>
