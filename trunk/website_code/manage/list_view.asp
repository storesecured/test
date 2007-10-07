<!-- start of list view -->
<%

if Service_Type < myStructure("Level") then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		<% if myStructure("Level") = 3 then %>
			BRONZE
		<% elseif myStructure("Level") = 5 then %>
			SILVER
		<% elseif myStructure("Level") = 7 then %>
			GOLD
		<% elseif myStructure("Level") = 9 then %>
			PLATINUM
		<% elseif myStructure("Level") = 11 then %>
			UNLIMITED
		<% end if %>
		Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>

<% else %>
	<SCRIPT LANGUAGE="Javascript">

	<!--
	function VerifyDelete(OID)
	{
		return confirm("Are you sure you want to delete the selected <%= myStructure("CommonName") %>?");}
	// -->
	
	function surfto(form) {
    var myindex=form.page.selectedIndex
    if (form.page.options[myindex].value != '0') {
    location=form.page.options[myindex].value;}}

	function selectAllItems(){
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
				document.forms[0].DELETE_IDS[i].checked = document.forms[0].selectall.checked; }
				
	function goDeleteSeleted(){
		tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to delete.");
                }
                 else
                 {
		if (confirm("Are you sure you want to delete the selected <%= myStructure("CommonName") %>s?")){
			document.forms[0].Delete_Id.value="SEL";
			document.forms[0].submit(); }}}

		function selectAllItems(){
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
				document.forms[0].DELETE_IDS[i].checked = document.forms[0].selectall.checked; }

	</script>
	<input type="Hidden" name="Delete_Id" value="">
	<% if myStructure("BackTo") <> "" then %>
		<TR bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="15">&nbsp;
		<%
                if myStructure("BackToName")<>"" then
		      BackToName=myStructure("BackToName")
		else
		      BackToName=myStructure("CommonName")
		end if 
                %>
		<input type="button" OnClick='JavaScript:self.location="<%=myStructure("BackTo") %>"' class="Buttons" value="<%= BackToName %> List" name="Back"></td>
		</tr>
	<% end if

	iColumns = 0
		   
	if noRecords = 0 then
		sHead="<tr align='center' bgcolor='#0069B2' class='white'>"
		for each sColumn in split(myStructure("HeaderList"),",")
			col_head=myStructure("Heading:"&sColumn)
			col_name=sColumn
			if myStructure("PrimaryKey") <> col_name or myStructure("Heading:"&col_name)<>"PK"  then
				iColumns = iColumns + 1
				sHead = sHead & "<td class=tablehead>"
				iSortable=0
				for each sColumnSort in split(myStructure("SortList"),",")
					if sColumn=sColumnSort then
						iSortable=1
						exit for
					end if
				next
				iFound = 0
				if Sort_By_Compare = "" then
					Sort_By_Compare = Sort_By
				end if
      			iFound = 1
      			if iFound = 1 and iSortable=1 then
  					if col_name = Sort_By_Compare then
						if sort_direction = 1 then
							sHead = sHead & "<B><a class=white href="&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(col_name)&"&Sort_Direction=0&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&">"&col_head&"<img src=images/desc.gif border=0></a></b>"
						else
							sHead = sHead & "<B><a class=white href="&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(col_name)&"&Sort_Direction=1&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&">"&col_head&"<img src=images/asc.gif border=0></a></b>"
						end if
  					else
						if col_head <> "None" then
							sHead = sHead & "<B><a class=white href="&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(col_name)&"&Sort_Direction=0&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&">"&col_head&"<img src=images/spacer.gif border=0 width=13></a></b>"
						else
							sHead = sHead & "&nbsp;"
						end if
  					end if
  				else
  					'no sorting on this column
					sHead = sHead & "<font class=white>"&col_head&"</font>"
  				end if
  				sHead = sHead & "</td>"
			end if
		next
		if myStructure("EditAllowed") then
			sHead = sHead & "<TD class=tablehead><font class=white>Edit</font></td>"
		end if
		if myStructure("DeleteAllowed") then
		   sHead = sHead & "<TD class=tablehead><font class=white>Delete</font></TD><TD class=tablehead>&nbsp;</td>"
		end if
		sHead = sHead & "</tr>"

		sRecordsPerPage= myStructure("CommonName")&"s per page <select name=page onChange='surfto(this.form)'>"
		if RowsPerPage = 25 then
		   selected25="selected"
		elseif RowsPerPage = 50 then
		   selected50="selected"
		elseif RowsPerPage = 100 then
		   selected100="selected"
		elseif RowsPerPage = 250 then
		   selected100="selected"
		elseif RowsPerPage = 500 then
		   selected500="selected"
		end if
		sRecordsPerPage=sRecordsPerPage&  "<option "&selected25&" value='"&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(Sort_By)&"&Form="&server.urlencode(myStructure("Form"))&"&Records=25'>25</option>"
		sRecordsPerPage=sRecordsPerPage&  "<option "&selected50&" value='"&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(Sort_By)&"&Form="&server.urlencode(myStructure("Form"))&"&Records=50'>50</option>"
		sRecordsPerPage=sRecordsPerPage&  "<option "&selected100&" value='"&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(Sort_By)&"&Form="&server.urlencode(myStructure("Form"))&"&Records=100'>100</option>"
		sRecordsPerPage=sRecordsPerPage&  "<option "&selected250&" value='"&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(Sort_By)&"&Form="&server.urlencode(myStructure("Form"))&"&Records=250'>250</option>"
		sRecordsPerPage=sRecordsPerPage&  "<option "&selected500&" value='"&myStructure("FileName")&"Start_Row="&Start_Row&"&Sort_By="&server.urlencode(Sort_By)&"&Form="&server.urlencode(myStructure("Form"))&"&Records=500'>500</option>"
		sRecordsPerPage=sRecordsPerPage&  "</select>"

		sRecordNav = "<TR bgcolor='#FFFFFF'><td width='100%'><table width='100%' border='0' cellspacing='1' cellpadding=2><tr bgcolor='#FFFFFF'>"
		if Start_Row > 0 then
			sRecordNav = sRecordNav & "<td width='33%'><a class=link href="&myStructure("FileName")&"Start_Row="&(Start_Row - RowsPerPage)&"&Sort_By="&Sort_By&"&Sort_Direction="&Sort_Direction&"&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&"><< Prev "&RowsPerPage&"</a></td>"
  		else
			sRecordNav = sRecordNav & "<td width='33%'>&nbsp;</td>"
		end if
		sRecordNav = sRecordNav & "<td width='33%' align=center nowrap>Displaying "&(Start_Row+1)&" - "&(End_Row+1)&" of "&(myfields("rowcount")+1)&"</td>"
		if myfields("rowcount") > End_Row then
			sRecordNav = sRecordNav & "<td width='33%' align=right><a class=link href="&myStructure("FileName")&"Start_Row="&(Start_Row + RowsPerPage)&"&Sort_By="&Sort_By&"&Sort_Direction="&Sort_Direction&"&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&">Next "&RowsPerPage&" >></a></td>"
  		else
			sRecordNav = sRecordNav & "<td width='33%'>&nbsp;</td>"
		end if
		sRecordNav = sRecordNav & "</tr><tr bgcolor='#FFFFFF'>"

		iPage=1
		sRecordNav = sRecordNav & "<td colspan=2 align=left>"
		if formatnumber(myfields("rowcount"),0)>formatnumber(rowsperpage,0) then
			sRecordNav = sRecordNav & "Jump to Page " 
			FOR rowcounter= 0 TO myfields("rowcount")/rowsperpage
				if rowcounter<>(myfields("rowcount")/rowsperpage) and rowcounter<>0 and (rowcounter*rowsperpage)>(Start_Row-(rowsperpage*7)) and (rowcounter*rowsperpage)<=(Start_Row+(rowsperpage*7)) then
					sRecordNav=sRecordNav&"&nbsp;|&nbsp;"
				end if
				if (rowcounter*rowsperpage)>=(Start_Row-(rowsperpage*6)) and (rowcounter*rowsperpage)<=(Start_Row+(rowsperpage*6)) then
					if Start_Row=rowcounter*rowsperpage then
					  sRecordNav = sRecordNav & "<B>" & iPage & "</b>"
					else
					  sRecordNav = sRecordNav & "<a class=link href="&myStructure("FileName")&"Start_Row="&((iPage-1)*RowsPerPage)&"&Sort_By="&Sort_By&"&Sort_Direction="&Sort_Direction&"&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&">"&iPage & "</a>"
					end if
				elseif (rowcounter*rowsperpage)=(Start_Row-(rowsperpage*7)) then
					  sRecordNav = sRecordNav & "<a class=link href="&myStructure("FileName")&"Start_Row="&((iPage-1)*RowsPerPage)&"&Sort_By="&Sort_By&"&Sort_Direction="&Sort_Direction&"&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&"><<</a>"
				elseif (rowcounter*rowsperpage)=(Start_Row+(rowsperpage*7)) then
					  sRecordNav = sRecordNav & "<a class=link href="&myStructure("FileName")&"Start_Row="&((iPage-1)*RowsPerPage)&"&Sort_By="&Sort_By&"&Sort_Direction="&Sort_Direction&"&Form="&server.urlencode(myStructure("Form"))&"&Records="&RowsPerPage&">>></a>"
				end if
				iPage=iPage+1
				
			next
		end if
		sRecordNav = sRecordNav & "</td>"
		
		response.write sRecordNav & "<td align=right>"&sRecordsPerPage&"</td></tr></table></td></tr><TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='13'><table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>" & sHead

		if myfields("rowcount") < End_Row then
			End_Row = myfields("rowcount")
		end if	

		str_class=1
		FOR rowcounter= 0 TO (End_Row-Start_Row)
			thisId = mydata(myfields(myStructure("PrimaryKey")),rowcounter)

			if str_class=1 then
				str_class=0
				bgcolor="#FFFFFF"
			else
				str_class=1
				bgcolor="#ECEFF0"
			end if
				response.write "<tr class='"&str_class&"' bgcolor='"&bgcolor&"'>"
				for each sColumn in split(myStructure("HeaderList"),",")
					col_head=myStructure("Heading:"&sColumn)
					col_format=myStructure("Format:"&sColumn)
					col_link=""
					col_link=myStructure("Link:"&sColumn)

					if myStructure("PrimaryKey") <> sColumn or myStructure("Heading:"&sColumn)<>"PK" then
						sDataBefore = mydata(myfields(sColumn),rowcounter)
						if col_format="INT" then
							iDecimalPlaces = myStructure("Length:"&sColumn)
							if iDecimalPlaces <> "" then
								sData = FormatNumber(sDataBefore,iDecimalPlaces)
							else
								sData = sDataBefore
							end if
						elseif col_format = "STRING" then
							iLength = myStructure("Length:"&sColumn)
							if iLength <> "" then
								sData = left(sDataBefore,iLength)
							else
								sData = sDataBefore
							end if
							sFileName = sData
						elseif col_format = "TEXT" then
							sData = col_head
							
						elseif col_format = "DATE" then
							 sData=formatdate(sDataBefore,"%M/%D/%y %H:%N")
						elseif col_format = "DATESHORT" then
							sData=formatdate(sDataBefore,"%M/%D/%y")
						elseif col_format = "LOOKUP" then
							sLookup = myStructure("Lookup:"&sColumn)     
							for each sOption in split(sLookup,"^")
								sOptionArray = split(sOption,":")
								if sOptionArray(0) = sDataBefore or (sOptionArray(0)="True" and sDataBefore) or (sOptionArray(0)="False" and not(sDataBefore)) or (isNumeric(sOptionArray(0)) and isNumeric(sDataBefore) and formatnumber(sOptionArray(0),0)=formatnumber(sDataBefore,0)) then
									sData = sOptionArray(1)
								end if
							next
							sData = sData
						elseif col_format = "CURR" then
							sData = Currency_Format_Function(sDataBefore)
						elseif col_format = "PERCENT" then
							sData = sDataBefore &"%"
						elseif col_format = "SQL" then
							sSQL = myStructure("Sql:"&sColumn)
							sSQL = replace(sSQL,"PK",thisId)
							sSQL = replace(sSQL,"THISFIELD",sDataBefore)  
							Set rs_store1 = Server.CreateObject("ADODB.Recordset")
							session("sql")=sSQL
							rs_store1.open sSQL,conn_store,1,1
							if not rs_store1.eof and not rs_store1.bof then
							   sData=rs_store1(sColumn)
							else
							   sData=myStructure("Default:"&sColumn)
							end if
							rs_store1.close
							set rs_store1 = nothing
						elseif col_format = "COLLECTION" then
							sCollectionArray = split(myStructure("Collection:"&sColumn),",")
							sCollectionName = sCollectionArray(1)
							sCollectionId = sCollectionArray(0)
							sData = "<select name='collection' size='1'>"
							Collection = sDataBefore
							sFound=0
							if noRecords1 = 0 then
								FOR rowcounter1= 0 TO myfields1("rowcount")
									one_item = Cstr(mydata1(myfields1(sCollectionId),rowcounter1))
									' diplay only those values that are inside our collection list
									' let's use this function :	Is_In_Collection(Collection,one_item,Del)
									if Is_In_Collection(Collection,one_item,",") then
										  sData = sData & "<Option value='"&mydata1(myfields1(sCollectionName),rowcounter1)&"'>"&mydata1(myfields1(sCollectionName),rowcounter1)
										sFound=1
									end if
								Next
							end if
							if sFound=0 then
							   sData = sData & "<Option value='"&myStructure("Default:"&sColumn)&"'>"&myStructure("Default:"&sColumn)
							end if
							sData = sData & "</select>"
						end if
					  
						if col_link<>"" then
							col_link = replace(col_link,"PK",thisId)
							col_link = replace(col_link,"THISFIELD",sDataBefore)
							col_link = replace(col_link,"THISURL",myStructure("FileName"))
							sData = "<a href="&col_link&" class=link>"&sData&"</a>"
						end if

						if sData="" then
							sData="&nbsp;"
						end if
						response.write "<TD class='"&str_class&"'>"&sData&"</td>"
					end if
				next
				if myStructure("EditAllowed") then %>
					<td nowrap class='<%=str_class%>'><a class="link" href="<%= myStructure("EditRecord") %>op=edit&Id=<%= thisId %>">Edit</a></td>
				<% end if

				if myStructure("DeleteAllowed") then %>
					<td nowrap class='<%=str_class%>'><a class="link" href="<%= myStructure("FileName") %>Start_Row=<%= Start_Row %>&Sort_By=<%= Sort_by %>&Form=<%= server.urlencode(myStructure("Form"))%>&Delete_Id=<%= thisId %>&Records="&RowsPerPage&" OnClick="return VerifyDelete('0');">Delete</a></td>
					<td align=center class='<%=str_class%>'><input class="image" type="checkbox" name="DELETE_IDS" value="<%= thisId %>"><input type="hidden" name="DELETE_FILENAMES" class="image" value="<%=sFileName%>">
				<% end if %>
				</tr>
		<% Next %>
		<% if myStructure("EditAllowed") then
			  iColumns = iColumns + 1
		end if
		if myStructure("DeleteAllowed") then
			  iColumns = iColumns + 1
		end if %>
		<tr class=0 bgcolor='#FFFFFF'><td colspan=<%= iColumns %> align=right>
		<% if myStructure("Footer") <> "" then
			  response.write myStructure("Footer")
		end if
		if myStructure("AddAllowed") then %>
			<input type="button" OnClick=JavaScript:self.location="<%= myStructure("NewRecord") %>" class="Buttons" value="Add <%=myStructure("CommonName") %>" name="Create_new">
		<% end if
		if myStructure("DeleteAllowed") then %>
			<input class=buttons type="Button" name="DeleteAll" value="Delete Selected" onClick="JavaScript:goDeleteSeleted();"></td><td align=center>
			<% if End_Row-Start_Row>=1 then %>
				<input class="image" type="checkbox" name="selectall" onClick="JavaScript:selectAllItems();"><br>Check All</td>
			<% end if %>
		<% end if %>
		</tr>
		<% response.write sHead & "</table></td></tr>" & sRecordNav&"<td align=right></td></tr></table></td></tr>"
	Else
		response.write "<TR class=0 bgcolor='#FFFFFF'><TD>No "&myStructure("CommonName")&"s were found</td></tr>"
                if myStructure("AddAllowed") then
			response.write "<TR class=0 bgcolor='#FFFFFF'><TD><input type='button' OnClick=JavaScript:self.location="""&myStructure("NewRecord")&""" class='Buttons' value='Add "&myStructure("CommonName")&"' name='Create_new_Coupon'></td></tr>"
		end if
	End If
end if
set myfields = nothing

%>
<!-- end of list view -->
