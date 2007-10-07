<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
'sQuestion_Path = "advanced/custom_fields.htm"
sInstructions="A form can be used to collect specific information from a customer and have it emailed."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_forms"
myStructure("TableWhere") = ""
myStructure("ColumnList") = "fldpbid,fldpagename,fldsubjectform,fldtoemail, 'List' as fields"
myStructure("HeaderList") = "fldpbid,fldpagename,fldsubjectform,fldtoemail,fields"

myStructure("DefaultSort") = "fldpagename"
myStructure("PrimaryKey") = "fldpbid"
myStructure("Level") = 1
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "design"
myStructure("Title") = "Forms"
myStructure("FullTitle") = "Design > Forms"
myStructure("CommonName") = "Form"
myStructure("NewRecord") = "custom_fields_page.asp"
myStructure("EditRecord") = "custom_fields_page.asp"
myStructure("Level") = 0
myStructure("Heading:fldpbid") = "PK"
myStructure("Heading:fldpagename") = "Page Name"
myStructure("Heading:fldsubjectform") = "Subject"
myStructure("Heading:fldtoemail") = "To Email"
myStructure("Heading:fields") = "Fields"

myStructure("Format:fldpbid") = "INT"
myStructure("Format:fldpagename") = "STRING"
myStructure("Format:fldsubjectform") = "STRING"
myStructure("Format:fldtoemail") = "STRING"
myStructure("Format:fields") = "STRING"

myStructure("Link:fields") = "form_fields_list.asp?fldPBID=PK"

	if Request.QueryString("Delete_Id") <> "" then
		

	    delete_id = Request.QueryString("Delete_Id")

		sql_Delete_Details = "delete from Store_form_fields where fldPBID =" & delete_id & " and store_id = " & store_id
		conn_store.execute sql_Delete_Details
			

			sql_pageID = "select fldPageID from store_forms where fldPBID = " & delete_id & " and store_id = " & store_ID


		set rs_pageID  = server.createobject("ADODB.Recordset")
		
			rs_PageID.open sql_pageID, conn_store,1,1
			if not rs_pageID.eof then
				page_ID =rs_pageID.fields("fldPageID")
			end if
			rs_pageID.close
		
		set rs_pageID =  nothing

		update_SQL =  "Update store_pages  set  page_form_content ='' where page_id = " & page_id & " and store_id = " & store_id
		conn_store.execute update_SQL

	end if


'MULTIPLE DELETE

	if  Request.Form("Delete_Id") ="SEL" then
  		if request.form("DELETE_IDS") <> "" then
			delete_id = request.form("DELETE_IDS")
	
			sql_Delete_Details = "delete from pageBuilder_Items where fldPBID in (" & delete_id & ") and store_id = " & store_id
			conn_store.execute sql_Delete_Details


			sql_pageID = "select fldPageID from pageBuilder where fldPBID in (" & delete_id & ") and store_id = " & store_ID
			response.write "<br><br><br><br><h4>"&sql_pageID  &"</h4>"
			set rs_pageID  = server.createobject("ADODB.Recordset")
		
			rs_PageID.open sql_pageID, conn_store,1,1
			response.write "Recordset opened<br>"
			if not rs_pageID.eof then	 
				response.write "Inside IF Not EOF<br>"
				while not rs_pageID.eof 
					page_ID = page_id & rs_pageID.fields("fldPageID") &  ","
						
					rs_pageID.movenext
				wend
			end if
			rs_pageID.close
			set rs_pageID =  nothing
			page_id = trim(page_id)
			page_id =   left(page_id,len(page_id)-1)

			update_SQL =  "Update store_pages  set  page_form_content ='' where page_id in (" & page_id & ") and store_id = " & store_id
			response.write "<br>" & update_Sql
			conn_store.execute update_SQL
			
		end if
    end if
	%>

<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->


<%
	set myStructure=nothing
	createFoot thisRedirect, 0%>
 
