<!--#include file="Global_Settings.asp"-->
<!--#include file="createForm.asp"-->
<% 
if not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if
'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%> <!--#include virtual="common/Error_Template.asp"--><% 
else



			Custom_Field_Type = Request.Form("Custom_Field_Type")
			varpageid_u= Request.Form("CustomField_ID")

			if Custom_Field_Type = 1 then
				Custom_Field_Values = Request.Form("Custom_Text_Value")
				if not isNumeric(Custom_Field_Values) then
					fn_error "Please enter the size of the text field."
				end if
			elseif Custom_Field_Type = 2 then
				Custom_Field_Values = Request.Form("Custom_Text_Area_Value")
				if not isNumeric(Custom_Field_Values) then
					fn_error "Please enter the number of columns for the textarea."
				end if
			elseif  Custom_Field_Type = 3 then
				Custom_Field_Values = checkStringForQ(Request.Form("Custom_Combo_Value"))
			elseif  Custom_Field_Type = 4 then
				Custom_Field_Values = checkStringForQ(Request.Form("Custom_Check_Value"))
			else
			    fn_error "Please choose the type of field you want to use."
                        end if
			if Request.Form("Field_Required") <> "" then
			    Field_Required = 1
			else
			    Field_Required = 0
			end if
			Custom_Field_Name=checkstringforQ(Request.Form("Custom_Field_Name"))
			View_Order=Request.Form("View_Order")
			if Request.QueryString("id")<>"" then 
				Editid=Request.QueryString("id") 
			else
				Editid=Request.Form("id") 
			end if
			
			op=Request.form("op") 
			'Response.Write  "OP====" & op & " ======= " & varpageid_u 
			'Response.End
			if Editid<>"" or op="edit" then
				strupdate=" update Store_form_fields set fldname='"&Custom_Field_Name&"' ,FldRequired="&Field_Required&",fldViewOrder="&View_Order&",fldType='"&Custom_Field_Values&"',fldFieldtype=" &Custom_Field_Type&" where fldid=" &varpageid_u
				'strinsert=strinsert & " values ( getdate()," & store_id & ","&varpageid_u&",'" & Custom_Field_Name & "','" &Field_Required& "','" &View_Order& "','"&Custom_Field_Values&"')"
				'Response.Write "<br> strupdate=" & strupdate
				session("sql")=strupdate
                                conn_store.Execute strupdate
				'Response.End 
				
				set rs_pageId =  server.createObject("ADODB.Recordset")

				rs_pageId.open "select fldPbID from store_form_fields where fldID = " & varpageid_u, conn_store, 1,1
					page_id =  rs_pageId.fields("fldPBID")
					page_id1 = page_id 
				rs_pageId.close
				
				set rs_pageId = nothing
				
				call generateForm(page_id)	
				'response.write "<br>select fldPbID from pageBuilder_items where fldID = " & varpageid_u &"<br>"
				Response.Redirect "form_fields_list.asp?fldpbid="& page_id1
			
			else
			
			strinsert=" insert into Store_Form_fields (Store_ID,fldpbid,fldname,FldRequired,fldViewOrder,fldType,fldFieldtype) "
			strinsert=strinsert & " values ("&store_id & ","&varpageid_u&",'" & Custom_Field_Name & "'," &Field_Required& "," &View_Order& ",'"&Custom_Field_Values&"',"&Custom_Field_Type&")"
			session("sql")=strinsert
			conn_store.Execute strinsert
			'Response.End 
			
			call generateForm(varpageid_u)
			'page_id = varpageid_u
			Response.Redirect "form_fields_list.asp?fldpbid="&varpageid_u
			end if
			'
			


end if 

%>
