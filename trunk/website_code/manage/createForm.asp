<%
'1-text box  2- textarea 3 - combo
	function GenerateForm(pageId)
		set rs = server.createobject("ADODB.Recordset")
		set rsForm  = server.createobject("ADODB.Recordset")
		sqlPb1 = "select * from store_forms where store_id = " & store_id & " and FldPBId =" & pageID
		'response.write sqlPb1
		rs.open sqlPb1 , conn_store,1,1
		if not rs.eof then
'			HTMLString = HTMLString +  "<form name=frmCustomEmail action= form_email.asp method=post ><table><tr><td><b>" & rs("FldSubjectForm")  & "</b></td><td></td></tr><tr><td><input type=hidden name=Send_To  value = "& rs("fldToEmail") & "></td><td></td></tr><tr><td><input type=hidden name=pageName value=" & rs("fldPageName") & "><input type=hidden name=pageId value=" & rs("fldPageId") & "><input type=hidden name=PBId value=" & rs("fldPBId") & "></td><td></td></tr>"

					HTMLString = HTMLString +  "<form name=frmCustomEmail action= form_email.asp method=post ><table><tr><td><b>" & rs("FldSubjectForm")  & "</b></td><td></td></tr><tr><td><input type=hidden name=Send_To  value = "& rs("fldToEmail") & "></td><td></td></tr>"

			SqlPbFields = "select * from store_form_fields where store_id = " & store_id & " and fldPBID = " & pageID & " order by FldViewOrder"
	
			rsForm.open SqlPbFields,conn_store,1,1
			if rsForm.eof then
				HTMLString = ""
			else
		
				while not rsForm.eof
						if rsForm("fldfieldtype") = 1 then
							HTMLString = HTMLString +  "<tr><td>"& rsForm("fldName") &"</td><td><input type=text name='" & rsForm("fldName") &"' size= " & rsForm("fldType") & ">"
							if rsForm("fldRequired") then
								HTMLString =  HTMLString + "<INPUT type=hidden value='Re|String|0|200|||" & rsForm("fldName") & "' name='"& rsForm("fldName") &"_C'>*"
							else
								HTMLString =  HTMLString +"<INPUT type=hidden value='Op|String|0|200|||"& rsForm("fldName") &"' name='" & rsForm("fldName") &"_C'>"
							end if
							response.write "</td></tr>"

						elseif rsForm("fldfieldtype") =2 then
							HTMLString = HTMLString +  "<tr><td>" & rsForm("fldName") &"</td><td><textarea name='" & rsForm("fldName") &"' cols = " & rsForm("fldType") & "></textarea>"
							if rsForm("fldRequired") then
								HTMLString =  HTMLString + "<INPUT type=hidden value='Re|String|0|200|||" & rsForm("fldName") & "' name='"& rsForm("fldName") &"_C'>*"
							else
								HTMLString =  HTMLString +"<INPUT type=hidden value='Op|String|0|200|||"& rsForm("fldName") &"' name='" & rsForm("fldName") &"_C'>"
							end if
							response.write "</td></tr>"

						elseif rsForm("fldfieldtype") =3 then
							HTMLString = HTMLString +  "<tr><td>" & rsForm("fldName") &"</td><td><select name='" & rsForm("fldName") & "'>"
								''Split the values into an array
								valueArray  = split(rsForm("fldType"), ",")
								for i=0 to ubound(valueArray)

									HTMLString = HTMLString +  "<option value = '" &  valueArray(i) & "'> " &  valueArray(i) &"</option>"
								next
							HTMLString = HTMLString +  "</select>"
                                                        if rsForm("fldRequired") then
								HTMLString =  HTMLString + "<INPUT type=hidden value='Re|String|0|200|||" & rsForm("fldName") & "' name='"& rsForm("fldName") &"_C'>*"
							end if
                                                        HTMLString =  HTMLString + "</td></tr>"
	
						else
								HTMLString = HTMLString +  "<tr><td>" & rsForm("fldName") &"</td><td>"
								''Split the values into an array
								valueArray  = split(rsForm("fldType"), ",")
								HTMLString = HTMLString +  "<input type=checkbox name='" & rsForm("fldName") &"' value='check'> "
                        if rsForm("fldRequired") then
   								HTMLString =  HTMLString + "<INPUT type=hidden name='" & rsForm("fldName") & "' value=' '>"
   								HTMLString =  HTMLString + "<INPUT type=hidden value='Re|String|0|200|check||" & rsForm("fldName") & "' name='"& rsForm("fldName") &"_C'>*"
   							end if
							HTMLString = HTMLString +  "</td></tr>"


						end if
					rsForm.movenext
				wend
				HTMLString =HTMLString +"<tr><td></td><td><input type=submit  value=Submit ></td></tr></table></form>"

			
			end if

		
		'	response.write "<br>"
		'	response.write  htmlsTRING
		'	response.write "<br>"
			SqlFormInsert = "update store_pages set page_form_content = '" & nullifyQ(HTMLString) &"' where store_id=" & store_id & " and page_id = " & rs("fldpageId")

                        page_id = rs("fldpageId")
			conn_store.execute SqlFormInsert
			HTMLString=""
		end if
	set rs = nothing
	set rsForm = nothing

end function
	
	
	'call GenerateForm
%>
