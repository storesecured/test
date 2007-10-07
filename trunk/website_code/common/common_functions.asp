<%
function fn_check_wait(sSeconds)
	'first make sure user is still there
	fn_print_debug "Checking if user is still connected"
	if response.isclientconnected=false then
		fn_error "Client Disconnected"
	end if

	if isNumeric(sSeconds) and sSeconds<>0 then
		fn_print_debug "Waiting "&sSeconds
		fn_wait(sSeconds)
	end if
end function

function fn_page_url (sFilename, IsLink)
    if IsLink<>0 then
        sLink = sFilename
    elseif sFilename = "store.asp" then
        sLink = Switch_Name
    elseif sFilename = "browse_dept_items.asp" then
        sLink = fn_dept_url("","")
    else
        sLink = Switch_Name&sFilename
    end if
    fn_page_url = sLink

end function

function fn_item_url (sFullName,sItemName)
    if sFullName<>"" or sFullName="Un Assigned" then
        sDeptString = fn_transform_fullname(sFullName)&"/"
    end if

    if sItemName<>"" then
        fn_item_url = Switch_Name&"items/"&sDeptString&sItemName&"-detail.htm"   
    else
        fn_item_url = Switch_Name&"items/"&sDeptString&"list.htm"  
    end if
end function

function fn_dept_url (sFullName,sJumptoPage)
    if sFullName<>"" or sFullName="Un Assigned" then
        sDeptString = fn_transform_fullname(sFullName)&"/"
    else
        sDeptString = ""
    end if
    
    if sJumptoPage="0" then
        sJumptoPage=""
    end if  
    
    fn_dept_url = Switch_Name&"items/"&sDeptString&"list"&sJumptoPage&".htm"
end function

function fn_redirect (sUrl)
    fn_print_debug "would temp redirect here to new <a href='"&sUrl&"'>"&sUrl&"</a>"
    if fn_is_debug=1 and 1=1 then
		response.end
	end if
	response.redirect sUrl
end function

function fn_redirect_perm (sUrl)
    fn_print_debug "would redirect here to new "&sUrl
    if fn_is_debug=1 then
		response.end
	end if
	response.status ="301 Moved Permanently"
	response.addheader "Location", sUrl
	response.end
end function

function fn_replace (sString,sReplaceWhat,sReplaceWith)
	if instr(sString,sReplaceWhat) then  
		if isNull(sReplaceWith) then
			sReplaceWith=""
		end if
		fn_replace = replace(sString,sReplaceWhat,sReplaceWith)
	else
		fn_replace = sString
	end if
end function

function fn_is_debug ()
	if store_id=sDebugStore and iDebugOn=1 and Request.ServerVariables("REMOTE_ADDR")=sDebugIP then
		fn_is_debug=1
	else
	    fn_is_debug=0
	end if
end function

function fn_print_debug (sDebug)
	if fn_is_debug=1 then
		response.write "<BR>"&sDebug
	end if
end function

function fn_append_querystring(sUrlStart,sAddQString)
	if instr(sUrlStart,"?")>0 then
		sUrlEnd=sUrlStart&"&"&sAddQString
	else
		sUrlEnd=sUrlStart&"?"&sAddQString
	end if
     fn_append_querystring=sUrlEnd
end function

function fn_get_querystring (sQString)
	sQString=request.querystring(sQString)
     sQString=fn_sanitize(sQString)

     fn_get_querystring=sQString
end function

function fn_sanitize(sSanitize)
	sSanitize=fn_replace(sSanitize,">","")
     sSanitize=fn_replace(sSanitize,"<","")
     sSanitize=fn_replace(sSanitize,"'","")
     sSanitize=fn_replace(sSanitize,"(","")
     sSanitize=fn_replace(sSanitize,")","")
     
     fn_sanitize=sSanitize
end function

function fn_wait(iSeconds)
	if isNumeric(iSeconds) then
		Set WaitObj = Server.CreateObject ("WaitFor.Comp")
		WaitObj.WaitForSeconds iSeconds
		set WaitObj=Nothing
	end if
end function

function fn_append_list (sString,sDelimiter,sAppend)
    if sString="" then
        sString=sAppend
    else
        sString=sString&sDelimiter&sAppend
    end if
    fn_append_list=sString
end function

function fn_transform_fullname (sFullName)
    'take a full name dept and translate to querystring
    'fn_print_debug "in transform fullname "&sFullName
    sMyName=""
    for each sThisDeptName in split(sFullName," > ")
        	sThisDeptName=lcase(sThisDeptName)
          Set objRegExp = New RegExp
		objRegExp.Global = True
		objRegExp.IgnoreCase = True
		objRegExp.Pattern = "[^a-zA-Z0-9]"
		sThisDeptName = objRegExp.Replace(sThisDeptName,"-")
		sThisDeptName=fn_replace(sThisDeptName,"--------","~")
		sThisDeptName=fn_replace(sThisDeptName,"-------","~")
		sThisDeptName=fn_replace(sThisDeptName,"------","~")
		sThisDeptName=fn_replace(sThisDeptName,"-----","~")
          sThisDeptName=fn_replace(sThisDeptName,"----","~")
          sThisDeptName=fn_replace(sThisDeptName,"---","~")
          sThisDeptName=fn_replace(sThisDeptName,"--","~")

        if sMyName="" then
            sMyName = sThisDeptName
        else
            sMyName = sMyName & "/"&sThisDeptName
        end if
    next
    fn_transform_fullname = sMyName
end function

Function fn_IsValidEmail(sEMail)
    ' original by Brad Murray
    ' optimized by Rob Hofker, email: rob@eurocamp.nl, 
     '23 august 2000

    ' Disallowed characters
    sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< "

    ' Check that there is at least one '@'
	if InStr(sEMail, "@") <= 0 then
	   fn_IsValidEmail=0
	   exit function
	end if

    ' Check that there is at least one '.'
	if InStr(sEMail, ".") <= 0 then
    		fn_IsValidEmail=0
    		exit function
	end if

    ' and that the length is at least six (a@a.ca)
	if Len(sEMail) < 6 then
		fn_IsValidEmail=0
		exit function
	end if

    ' Check that there is only one '@'
	i = InStr(sEMail, "@")
	sTemp = Mid(sEMail, i + 1)
	if InStr(sTemp, "@") > 0 then
     	fn_IsValidEmail=0
     	exit function
     end if

    'extra checks
    ' AFTER '@' space is not allowed
	if InStr(sTemp, " ") > 0 then
    		fn_IsValidEmail=0
    		exit function
    	end if

    ' Check that there is one dot AFTER '@'
	if InStr(sTemp, ".") = 0 then
    		fn_IsValidEmail=0
    		exit function
	end if

    ' Check if there's a quote (")
	if InStr(sEMail, Chr(34)) > 0 then
		fn_IsValidEmail=0
		exit function
	end if
    
        
    ' Check if there's any other disallowed chars
    ' optimize a little if sEmail longer than sInvalidChars
    ' check the other way around
    If Len(sEMail) > Len(sInvalidChars) Then
        For i = 1 To Len(sInvalidChars)
            If InStr(sEMail, Mid(sInvalidChars, i, 1)) > 0 then
          	fn_IsValidEmail=0
          	exit function
            end if
        Next
    Else
        For i = 1 To Len(sEMail)
            If InStr(sInvalidChars, Mid(sEMail, i, 1)) > 0 then
          	fn_IsValidEmail=0
          	exit function
            end if
        Next
    End If

    ' extra check
    ' no two consecutive dots
	if InStr(sEMail, "..") > 0 then
    		fn_IsValidEmail=0
    		exit function
	end if
	
	fn_IsValidEmail=1

End Function

function fn_md5hash (sHashedText)
	Set mdhash = server.createobject("FusionHashing.MD5")
		fn_md5hash = mdhash.Hash(sHashedText)
	Set mdhash = Nothing
end function

' ================================================================
' CHECK IF AN ITEM IS IN A COLECTION
Function Is_In_Collection(Collection,one_item,Del)
	Is_In_Collection = False
	My_Collection_Array = Split(Collection,Del)
	' loop over array to find Collection , we will trim just in case we have spaces ...
	For each Collection_element in My_Collection_Array
		if trim(CStr(Collection_element)) = CStr(one_item) then 
			Is_In_Collection = True
		end if
	next
end function

' ================================================================
' SEND AN EMAIL
Sub Send_Mail(SentFrom,SendTo,Subject,Body)
	fn_print_debug "in send mail from="&sentfrom&", to="&sendTo&",subject="&subject&",body="&body
    	if Subject="" or isNull(Subject) then
    		Subject="No Subject"
    	end if
	SendTo_array = split(SendTo,",")
	for each one_SendTo in SendTo_array
		if one_SendTo <> "" then
			Set Mail = Server.CreateObject("Persits.MailSender")
			Mail.From = SentFrom
			Mail.AddAddress one_SendTo
			Mail.CharSet = "UTF-8"
            Mail.ContentTransferEncoding = "Quoted-Printable"
			Mail.Subject = Subject
			Mail.Body = Body
			Mail.Helo="mail.storesecured.com"
			Mail.Queue=True
			Mail.Send
		end if
	next

	Set Mail = Nothing
end sub

Sub Send_Mail_Html(SentFrom,SendTo,Subject,Body)
      if Html_Notifications or isNull(Html_Notifications) then
        'Body=replace(Body,vbcrlf,"<BR>")
		if Subject="" or isNull(Subject) then
			Subject="No Subject"
		end if
        Body="<html><head><base href='"&Site_name&"'>"&_
	   	"<style type='text/css'>.normal"&_
    		"{font-family:verdana,helvetica,arial,sans-serif;"&_
    		"font-size:10pt}"&_
		"</style></head><body>"&Body&"</body></html>"
		if SendTo<>"" then
  		  SendTo_array = split(SendTo,",")
  		  for each one_SendTo in SendTo_array
  			if one_SendTo <> "" then
            Set Mail = Server.CreateObject("Persits.MailSender")
    			Mail.From = SentFrom ' Specify sender's address
    			Mail.AddAddress one_SendTo ' Name is optional
    			Mail.Subject = Subject
    			Mail.Body = Body
    			Mail.IsHtml = True
    			Mail.Helo="mail.storesecured.com"
    			Mail.Queue=True
    			Mail.Send
  			end if
  		  next
  		  end if
  		else
  		  Call Send_Mail(SentFrom,SendTo,Subject,Body)
      end if

end sub

' ================================================================
function checkStringForQ(theString)
	tmpResult = theString
	if not isNull(theString) then
		tmpResult = replace(tmpResult,"""","&#34;")
		tmpResult = replace(tmpResult,"'","&#39;")
	end if
	checkStringForQ = tmpResult
end function


SUB DataGetRows(parmConn,parmSQL,byref parmArray,byref parmDict,noRecords)
	dim conntemp,rstemp,howmany,counter
	set conntemp=server.createobject("adodb.connection")
	conntemp.open parmConn
	session("sql_old") = session("sql")
   session("sql") = parmSQL
	set rstemp=conntemp.execute(parmSQL)
	If rstemp.eof then
		noRecords=1
		parmdict.add "rowcount", -1
		rstemp.close
		set rstemp=nothing
		conntemp.close
		set conntemp=nothing
		exit sub
	end if

	' Now lets fill up dictionary with field names/numbers
	howmany=rstemp.fields.count
	for counter=0 to howmany-1
		if not parmdict.Exists(lcase(rstemp(counter).name)) then
			parmdict.add lcase(rstemp(counter).name), counter
		end if
	next

	noRecords=0

	' Now lets grab all the records
	parmArray=rstemp.getrows()
	parmdict.add "rowcount", ubound(parmarray,2)
	parmdict.add "colcount", howmany-1
	rstemp.close
	set rstemp=nothing
	conntemp.close
	set conntemp=nothing
END SUB

'CHECK VALUES OF THE FIELDS FROM A FORM
Function Form_Error_Handler(My_Form)
	' This function gets a form collection and return an Error string , if not empty that 
	' means that we have errors

	' set the error to initial empty value :
	Error = ""
   on error resume next
   
   if err.number<>0 and request.form="" then
	 fn_error("The data you have tried to save has exceeded the amount that a single form field can hold.  Please try reducing the amount of data or splitting it between different fields.")
   end if
	For Each Key In My_Form 
		New_Key = Key&"_C"
		if My_Form(New_Key) <> "" then
			' that is the signal that we need to do validation for this form element ...
			' first split with | as delimiter 
			MyFormElementArray = split(My_Form(New_Key),"|",-1,1)
			'glosary of checks to perform :
			'position		value check 		error message ...
			'0 			Re 	required 	Field required
			'0 			Op 					Field Optional
			'1 (type)		String		  String
			'1 			Integer	Integer		value integer
			'2 min size (for text ..)			100					Field to long ...
			'3 max size (for text)
			'4 must have characters @,$,%
			'5 must NOT have characters @,$,%
			
			if UBound(MyFormElementArray) >= 6 then
				sName = MyFormElementArray(6)
			else
				 sName = Key
			end if

                        if sName<>"" then
			sText = LCase(My_Form(Key))
			if MyFormElementArray(0) = "Re" AND sText = "" then
				Error = Error & "Field <b>"& sName & "</b> is required | "
			end if

			' check to see if the type is ok ...	
			if MyFormElementArray(1) = "Integer" and not IsNumeric(sText) and sText <> "" then
				Error = Error & "Field <b>"& sName & "</b> must be a number | "
			end if
			
			if MyFormElementArray(1) = "Integer" and instr(sText,",")>0 then
				Error = Error & "Field <b>"& sName & "</b> cannot contain a comma | "
			end if
			
			if MyFormElementArray(1) = "Integer" and isNumeric(sText) then
            if sText>9999999 then
				   Error = Error & "Field <b>"& sName & "</b> cannot exceed 9999999 | "
			   end if
         end if

			if MyFormElementArray(1) = "Integer" and instr(sText,"$")>0 then
				Error = Error & "Field <b>"& sName & "</b> cannot contain a currency symbol."
			end if

			if MyFormElementArray(1) = "date" and not IsDate(sText) and sText <> "" then
				Error = Error & "Field <b>"& sName & "</b> must be a date | "
			end if
			if MyFormElementArray(1) = "date" and sText<>"" and IsDate(sText) then
            if DatePart("yyyy",sText)>2078 then
				   Error = Error & "Field <b>"& sName & "</b> must be a date before 1/1/2078 | "
			   end if
         end if
			If MyFormElementArray(1) = "Email" then
				if Instr(1, sText, "@") = 0 or	Instr(1, sText, ".") = 0 then
					Error = Error & "Field <b>" & sName & "</b> must be a valid email | "
				end if
				MyFormElementArray(1) = "String"
			end if

			' process the size check only if not null and required
			if (MyFormElementArray(2) <> "" or MyFormElementArray(3) <> "") and not(MyFormElementArray(0) = "Op" and len(sText) = 0) then
				' for strings only :
				if MyFormElementArray(1) = "String" then
					if len(checkstringforQ(sText)) > int(MyFormElementArray(3)) or len(sText) < int(MyFormElementArray(2)) then
						Error = Error & "Field <b>"& sName & "</b> must be between "&MyFormElementArray(2)&" and "&MyFormElementArray(3)&" characters, your current length is "&len(checkstringforQ(sText))&" characters | "
					end if
				end if
				' for numbers - Integers only :
				if MyFormElementArray(1) = "Integer" then
					if (sText) <> "" and isNumeric(sText) then
						if MyFormElementArray(3) = "" and MyFormElementArray(2) <> "" then
						if int(sText) < int(MyFormElementArray(2)) then
								Error = Error & "Field <b>"& sName & "</b> must be greater than or equal to "&MyFormElementArray(2) & "|"
							end if
						elseif MyFormElementArray(2) = "" and MyFormElementArray(3) <> "" then
						if int(sText) > int(MyFormElementArray(3)) then
								Error = Error & "Field <b>"& sName & "</b> must be less than or equal to "&MyFormElementArray(3) & "|"
							end if
						elseif int(sText) > int(MyFormElementArray(3)) or int(sText) < int(MyFormElementArray(2)) then
							Error = Error & "Field <b>"& sName & "</b> must be between "&MyFormElementArray(2)&" and "&MyFormElementArray(3)& "|"
						end if
					end if
				end if
			end if
			 ' end of process the size check only if not null


			 
			 if UBound(MyFormElementArray) >= 4 then
		 if MyFormElementArray(4) <> "" and (sText) <> "" then
			 My4Array = split(MyFormElementArray(4),",",-1,1)
			 For Each sString In My4Array
				 if instr(sText,sString) = 0 then
					 Error = Error & "Field <b>"&sName & "</b> must contain a '" & sString & "'|"
				 end if
			 next
		 end if
			 end if
			 
			 if UBound(MyFormElementArray) >= 5 then
		 if MyFormElementArray(5) <> "" and (sText) <> "" then
			 My5Array = split(MyFormElementArray(5),",",-1,1)
			 For Each sString In My5Array
				 if instr(sText,sString) <> 0 then
					 Error = Error & "Field <b>"&sName & "</b> cannot contain a '" & sString & "'|"
				 end if
			 next
		 end if
			 end if
			 
			 if MyFormElementArray(1) = "String" then
		 sBadWords = "xp_"
		 MyArray = split(sBadWords,",",-1,1)
		 For Each sString in MyArray
			  if instr(sText,sString) <> 0 then
			Error = Error & "Field <b>"&sName&"</b> cannot contain '"&sString&"'|"
			  end if
		 next
			 end if
                  end if
		end if
	Next

	Form_Error_Handler = Error

End Function

' ================================================================
' FORMAT A NUMBER AS A CURRENCY
' FUNCTION SUPORT INTERNATIONAL CURRENCIES CONVERSION
Function Currency_Format_Function (Number)
   if IsNull(Number) then
      Number = 0
   end if
   if isnumeric(Number) then
      if Store_Currency="" then
        this_Store_currency = Session("Store_currency")
      else
        this_Store_currency = Store_Currency
      end if
      Currency_Format_Function = this_Store_currency&FormatNumber(Number,2)
   else
      Currency_Format_Function = Number
   end if
End Function

function fn_get_group(this_store_id)
    fn_get_group= (this_store_id mod 29)
end function

function fn_get_sites_folder(this_store_id,sPath_type)
    Website_Folder=Application_Path&"customer_files\group_"&fn_get_group(this_store_id)&"\files_"&this_store_id&"\"
     
    if sPath_type="Base" then
        fn_get_sites_folder=Website_Folder
    elseif sPath_type="Root" then
        fn_get_sites_folder=Website_Folder&"website_root\"
    elseif sPath_type="Download" then
        fn_get_sites_folder=Website_Folder&"esd_download\"
    elseif sPath_type="Export" then
        fn_get_sites_folder=Website_Folder&"export\"
    elseif sPath_type="Key" then
        fn_get_sites_folder=Website_Folder&"keyfiles\"
    elseif sPath_type="Images" then
        fn_get_sites_folder=Website_Folder&"website_root\images\"
    elseif sPath_type="Log" then
        fn_get_sites_folder=Website_Folder&"raw_logs\"
    elseif sPath_type="Upload" then
        fn_get_sites_folder=Website_Folder&"upload\"
    elseif sPath_type="Statistics" then
        fn_get_sites_folder=Website_Folder&"statistics_data\"
    else
        fn_error "FilePath type "&sPath_type&" does not exist."
    end if
end function

function fn_get_code_folder(sPath_type)
    if sPath_type="Manage" then
        fn_get_code_folder=Application_Path & "website_code\manage\"
    elseif sPath_type="Store_Engine" then
        fn_get_code_folder=Application_Path & "website_code\store_engine\"
    elseif sPath_type="Common" then
        fn_get_code_folder=Application_Path & "website_code\common\"
    elseif sPath_type="Include" then
        fn_get_code_folder=Application_Path & "website_code\store_engine\include\"
    elseif sPath_type="Sales" then
        fn_get_code_folder=Application_Path & "website_code\sales\"
    elseif sPath_type="Logs" then
        fn_get_code_folder=Application_Path & "customer_files\logs\"
    elseif sPath_type="Attachments" then
        fn_get_code_folder=Application_Path & "auxillary_files\attachments\"
    elseif sPath_type="Certificates" then
        fn_get_code_folder=Application_Path & "auxillary_files\certificates\"
    elseif sPath_type="Images_Themes" then
        fn_get_code_folder=Application_Path & "auxillary_files\images_themes\"
    elseif sPath_type="Ftp" then
        fn_get_code_folder=Application_Path & "customer_files\ftp\"
    elseif sPath_type="Scheduled" then
        fn_get_code_folder=Application_Path & "auxillary_files\scheduled_scripts\"
    else
        fn_error "FilePath type "&sPath_type&" does not exist."
    end if
end function

function fn_error (sError)
	fn_redirect Switch_Name&"error.asp?Message_Id=101&Message_Add="&server.URLEncode(sError)
end function

function fn_url_rewrite (sUrl)
    if sUrl<>"" then
        sUrl = lcase(sUrl)
        sUrl = replace(sUrl,root_w_dot,sLocalAddName&root_w_dot)
        sUrl = replace(sUrl,".easystorecreator.com",sLocalAddName&".easystorecreator.com")
        sUrl = replace(sUrl,".easystorecreator.net",sLocalAddName&".easystorecreator.net")
        fn_url_rewrite = sUrl
    else
        fn_url_rewrite = ""
    end if
end function
%>