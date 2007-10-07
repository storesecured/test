<%

Server.ScriptTimeout =2000
Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")

sDelimiter = fn_get_delimiter()

File_Name = Request.Form("Import_Filename")
File_Full_Name = fn_read_excel(Upload_Folder,File_Name,sDelimiter)

fn_process_file arrColumns,File_Full_Name,sDelimiter,sProcName

function fn_read_excel(sPath, sFileName, sDelimiter)
	sFilePath=sPath&sFileName
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	ext=fs.GetExtensionName(sFilePath)
	if ext ="xls" Then
		on error resume next
		fn_print_debug "Found excel file "&sFilePath
		' Create a server connection object
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.Open "DBQ=" & sFilePath & ";" & _
		"DRIVER={Microsoft Excel Driver (*.xls)};"
		' Create a server recordset object

		if err.number=-2147024882 then
			call send_mail(sNoReply_email,sDeveloper_email,"Excel Problem",Server_IP&" not allowing connections")
			fn_error "There is a unknown error reading from your Excel file.  Please export it to a text file instead as the system is unable to read this file."
		end if
		rs_store.Open "SELECT * FROM [Sheet1$]", cn,3,2

		if err.number<>0 then
			rs_store.Close
			Set rs_store = Nothing
			cn.close
			set cs=Nothing
			' Kill the connection
			fn_error "Your Excel file must have its data in a sheet named Sheet1.  We were unable to find Sheet1 in the selected file.  Please reupload a file with the important worksheet renamed Sheet1."
		end if

		rs_store.movefirst
		dim fs,fname
		sFilePath=replace(sFilePath,".xls",".txt")
		set fname=fs.CreateTextFile( sFilePath,true)
		Do while not rs_store.eof
			datastore=""
			For counter = 0 to rs_store.fields.count - 1
				if  instr( rs_store.fields.item(counter).value ,chr(10)) > 0  then
					readmefull =""
					readmefull = replace(rs_store.fields.item(counter).value ,chr(13),"")
					readmefull = replace(rs_store.fields.item(counter).value ,chr(10),"")
					datastore=datastore & sDelimiter &  readmefull 
				else
					datastore=datastore & sDelimiter &  rs_store.fields.item(counter).value 
				end if
			Next 
			finaldata = right(datastore ,(len(datastore )-3))
			fname.WriteLine (finaldata)
			datastore=""
			rs_store.movenext
		Loop
		fn_print_debug "Wrote text file "&sFilePath

		fname.Close
		set fname=nothing
		set fs=nothing
		rs_store.Close
		Set rs_store = Nothing
		cn.close
		set cs=Nothing
		' Kill the connection
		
		if err.number<>0 then
			fn_error "The Excel driver does not support reading columns with a length of more then 255 chars.  Please ensure that each of your columns is less then 255 characters or export to a text file."&err.number&err.description
		end if
		fn_check_wait 0
	else
		if ext="jpg" or ext="gif" or ext="bmp" then
			fn_error "You cannot import information from a image file."
		end if
	END IF
	set fs=nothing
	fn_read_excel=sFilePath
	
end function

function fn_process_file(arrColumns,File_Full_Name,sDelimiter,sProcName)
	fn_print_debug "starting to process file"

     Set FileObject = CreateObject("Scripting.FileSystemObject")
	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 1)
	ReadAll = MyFile.ReadAll
	MyFile.close
	
	If instr(ReadAll,vbcrlf) <= 0 then
		FileObject.DeleteFile(File_Full_Name)
		ReadAll = replace(ReadAll,chr(13),vbcrlf)
		Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
		MyFile.Write ReadAll
		MyFile.Close
	End if
	Set MyFile = Nothing

	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 1)

	sError=""
	iLine=0
	on error resume next

	Do While MyFile.AtEndOfStream <> True
		ReadLineTextFile = MyFile.ReadLine
		if (iLine<=0) and instr(lcase(left(ReadLineTextFile,20)),lcase(sHeader)) and MyFile.AtEndOfStream <> True then
			'read next line, first line is header
			fn_print_debug "header line skipping "&iLine
			fn_print_debug "left 40="&left(ReadLineTextFile,20)
			fn_print_debug "sheader="&lcase(sHeader)
			ReadLineTextFile = MyFile.ReadLine
			iLine=iLine+1
		end if

		Line_Array = split(ReadLineTextFile,sDelimiter)
		sDataLine=""
		iColIndex=0

		isBlank=replace(ReadLineTextFile,sDelimiter,"")
		if iLine<>0 and isBlank="" then
			'we found a blank line so assume there isnt anything else
			fn_print_debug "Found blank line at "&iLine&isBlank
			exit do
		end if

		For Each sColumnDef in arrColumns
			sColumnDefArr=split(sColumnDef,":")
			sColumnName=sColumnDefArr(0)
			sColumnReq=sColumnDefArr(1)
			sColumnType=sColumnDefArr(2)
			if uBound(sColumnDefArr)>=3 then
				sColumnLength=cint(sColumnDefArr(3))
				if uBound(sColumnDefArr)>=4 then
					sColumnDefault=sColumnDefArr(4)
					if uBound(sColumnDefArr)>=5 then
						sColumnNotAllowed=sColumnDefArr(5)
					end if
				end if

			end if
			if uBound(Line_Array)>=iColIndex then
				sDataValue=Line_Array(iColIndex)
			else
				sDataValue=""
			end if
			


			if sColumnName="Country" then
				'CountryColumn = sColumnName
				CountryData = sDataValue
			end if
			
			if sColumnName="State" then
				'StateColumn = sColumnName
				StateData = sDataValue
			end if


			sDataValue = fn_remove_quotes(sDataValue)
			if sColumnType="Text" and sDataValue<>"" then
				sDataValue=checkstringforQ(sDataValue)
			elseif sColumnType="Html" and sDataValue<>"" then
				sDataValue=nullifyQ(sDataValue)
			end if

			if sDataValue="" and sColumnDefault<>"" then
				if instr(sColumnDefault,"%arrColumns%")=1 then
					iIndex=replace(sColumnDefault,"%arrColumns%","")
					sDataValue=Line_Array(iIndex)	
				else	
					sDataValue=sColumnDefault
				end if
			end if

			if sColumnReq="Re" and sDataValue="" then
				sError=sError&("<BR>"&sColumnName&" is required ("&sDataValue&")")
			elseif sColumnType="Date" and sDataValue<>"" and not isDate(sDataValue) then
				sError=sError&("<BR>"&sColumnName&" must be a date ("&sDataValue&")")
			elseif sColumnType="Integer" and sDataValue<>"" and not isNumeric(sDataValue) then
				sError=sError&("<BR>"&sColumnName&" must be numeric ("&sDataValue&")")
			elseif isNumeric(sColumnLength) and (sColumnType="Text" or sColumnType="Html") and (len(sDataValue)>sColumnLength) then
				sError=sError&("<BR>"&sColumnName&" maximum length "&sColumnLength&" ("&sDataValue&")")
			elseif sColumnType="Boolean" and sDataValue<>"" then
				if len(sDataValue)>1 or (ucase(sDataValue)<>"Y" and ucase(sDataValue)<>"N") then
					sError=sError&("<BR>"&sColumnName&" must be a Y or N ("&sDataValue&")")
				else
					if ucase(sDataValue)="Y" then
						sDataValue=1
					else
						sDataValue=0
					end if
				end if
			elseif sColumnNotAllowed<>"" then
				fn_print_debug "in not allowed "&sColumnNotAllowed
				For Each sNotString In split(sColumnNotAllowed,",")
	 				if instr(sDataValue,sNotString) <> 0 then
	 					sError=sError&("<BR>"&sColumnName&" cannot contain the following: ("&sNotString&")")
	 				end if
				next
			end if

			if sColumnType="Text" or sColumnType="Html" or sColumnType="Date" then
				sDataValue="'"&sDataValue&"'"
			elseif sDataValue="" and sColumnType="Integer" then
				sDataValue="null"
			end if
			
			if sColumnType="Integer" then
				sDataValue=cdbl(sDataValue)
			end if

			if (sDataValue="" or sDataValue="''") and sEmptytoNull=1 then
				sDataValue="Null"
			end if

			if iColIndex=0 then
				sDataLine=sDataValue
			else
				sDataLine=sDataLine&","&sDataValue
			end if
			
			iColIndex=iColIndex+1
			sColumnName=""
			sColumnReq=""
			sColumnType=""
			sColumnLength=""
			sColumnDefault=""
			sColumnNotAllowed=""
		next
		
		
		if CountryData = "United States" or CountryData = "Canada" then
			if len(StateData)>2 then
				Response.Redirect "admin_error.asp?Message_Add=In Case of United States and Canada, You can enter only 2 letters in STATE field"	    	
			'else
			end if
		end if
		

		if sError<>"" then
			fn_error_line arrColumns,sError,iLine+2,Line_Array
		end if

		err.number=0
		on error resume next

		sSqlStatement = sProcName&" "&store_id&","&sDataLine
		session("sql")=sSqlStatement
		fn_print_debug sSqlStatement
		conn_store.execute sSqlStatement
  		fn_print_debug sSqlStatement

		if err.number<>0 then
			fn_print_debug "In error="
			fn_print_debug "Error="&err.number
			fn_error_line arrColumns,err.description,iLine+2,Line_Array
		end if

		'if user is still connected, every 500 lines wait 5 seconds to throttle
		if iLine<>0 and iLine mod 500 = 0 then
			fn_check_wait 5
		end if
		iLine=iLine+1
	Loop

	fn_print_debug "Closing File "&sFilePath

	MyFile.close
	Set FileObject = Nothing
	Set MyFile = Nothing

	fn_print_debug "Done processing"
	fn_create_results iLine
end function

function fn_error_line (arrColumns,sError,iLine,Line_Array)
	fn_print_debug "in error line"

	sColumns=""
	iColIndex=0
	For Each sColumnDef in arrColumns
		sColumnDefArr=split(sColumnDef,":")
		sColumnName=sColumnDefArr(0)
		if uBound(Line_Array)>=iColIndex then
			sDataValue=Line_Array(iColIndex)
		else
			sDataValue=""
		end if
		sColumns=sColumns&"<BR>"&sColumnName&" = "&sDataValue
		iColIndex=iColIndex+1
	next

	fn_error "<B>Error Line "&iLine&": </b>"&sError&"<BR><BR><B>Line "&iLine&" Data</b>"&sColumns
end function

function fn_get_delimiter()
	'SELECT DELITER AND TEXT QUALIFIER
	
	select case cint(request.form("Delimiter"))
		case 0
			Delimiter="^{|"
		case 1
			Delimiter=","
		case 2
			Delimiter=";"
		case 3
			Delimiter=chr(9)
		case 4
			Delimiter=" "
		case 5
			Delimiter=request.form("Delimiter_Other")
	end select

	select case cint(Request.form("Qualifier"))
		case 1
			Delimiter = """"&Delimiter&""""
		case 2
			Delimiter = "''"&Delimiter&"''"
	end select

	fn_print_debug "Delimiter is "&Delimiter
	fn_get_delimiter=Delimiter
end function

function fn_remove_quotes (sQuotedText)
	sRight = right(sQuotedText,1)
	sLeft = left(sQuotedText,1)
	if sRight=sLeft then
		if sRight="""" or sRight="'" then
			sQuotedText = mid(sQuotedText,2,len(sQuotedText)-2)
			fn_print_debug "removing quotes="&sQuotedText
		end if
	end if
     fn_remove_quotes=sQuotedText
end function

function fn_file_delete(File_Full_Name)
	Set fs1=Server.CreateObject("Scripting.FileSystemObject") 
	if fs1.FileExists(File_Full_Name) =true  then
		  fn_print_debug "Deleting file "&File_Full_Name
		  fs1.DeleteFile(File_Full_Name )
	end if
	set fs1=nothing
end function

function fn_check_wait(sSeconds)
	'first make sure user is still there
	fn_print_debug "Checking if user is still connected"
	if response.isclientconnected=false then
		fn_error "Client Disconnected"
	end if

	if isNumeric(sSeconds) and sSeconds<>0 then
		fn_print_debug "Waiting "&sSeconds
		Set WaitObj = Server.CreateObject ("WaitFor.Comp")
		WaitObj.WaitForSeconds sSeconds
		set WaitObj=Nothing
	end if
end function

function fn_print_debug (sDebug)
	if store_id=217 and 1=1 then
		response.write "<BR>"&sDebug
	end if
end function

function fn_redirect(sRedirect)
	response.redirect sRedirect
end function

function fn_write(sRedirect)
	response.write sRedirect
end function

function fn_create_results(sUpdatedRows)
	createHead thisRedirect
	sText = "<tr bgcolor='#FFFFFF'>"&_
			"<td width='100%' height='1'><br>"&_
			"Import successful<br>"&_
			"Updated <b>"&sUpdatedRows&"</b> Rows"&_
			"</tr>"
	fn_write sText
	create_foot thisRedirect,0
end function



%>