<%

if Session("Department_List:"&Store_Id)="" then
	'build list up from db this one time so we dont need to later
	sql_select = "SELECT Department_ID, Department_Name, Full_Name, Last_Level from Store_Dept WITH (NOLOCK) where store_id = "&store_id&" and full_name<>'' order by Full_Name"
	set deptfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,deptdata,deptfields,noRecordsDept)
	sList=""
	if noRecordsDept = 0 then
		FOR deptrowcounter= 0 TO deptfields("rowcount")
			if deptdata(deptfields("last_level"),deptrowcounter) or deptdata(deptfields("department_id"),deptrowcounter)=0 then
				sName=deptdata(deptfields("full_name"),deptrowcounter)
			else
				sName=deptdata(deptfields("full_name"),deptrowcounter)
			end if
                        if len(sName)>70 then
                           sName=" . . . "&right(sName,67)
                        end if
			sList = sList & (deptdata(deptfields("department_id"),deptrowcounter)&"|{}|"&_
			sName&"|{new}|")
		next
	else
		sList="|{new}|"
	end if
	Session("Department_List:"&Store_Id)=sList

end if

function create_dept_list (sFieldName,sValue,sSize,sAddList)
	if sSize>1 then
		sSize=sSize&" multiple"
	end if

	sText="<select name='"&sFieldName&"' size="&sSize&">"
	sText = sText&"<option value=''>Choose One"

	sList=sAddList&Session("Department_List:"&Store_Id)
	for each sDepartmentRecord in split(sList,"|{new}|")
		if sDepartmentRecord<>"" then
			sDepartmentArray=split(sDepartmentRecord,"|{}|")
			sDepartment_Id=sDepartmentArray(0)
			sDepartment_Name=sDepartmentArray(1)
			sSelected = ""
			if (sValue<>"") then
				if Is_In_Collection(sValue,sDepartment_Id,",") then					
				'if formatnumber(sValue,0) = formatnumber(sDepartment_Id,0) then
				  sSelected = "selected"
				end if
			end if
			sText = sText&"<option "&sSelected&" value="&sDepartment_Id&">"&sDepartment_Name
		end if
	next
	sText=sText&"</select>"
	create_dept_list=sText

end function

%>