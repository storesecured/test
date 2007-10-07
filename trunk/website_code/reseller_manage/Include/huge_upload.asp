<%
  'Using Huge-ASP file upload
  Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
  'Using Pure-ASP file upload

  Server.ScriptTimeout = 2000
  Form.SizeLimit = &HA00000


  '{b}Set the upload ID for this form.
  'Progress bar window will receive the same ID.
  if len(Request.QueryString("UploadID"))>0 then
	  Form.UploadID = Request.QueryString("UploadID")'{/b}
  end if
  'was the Form successfully received?
  Const fsCompletted  = 0

  If Form.State = fsCompletted Then 'Completted
	 'was the Form successfully received?
	 if Form.State = 0 then
		'Do something with upload - save, enumerate, ...
    on error goto 0
    fUpload=1
    For Each File In Form.Files
       if sOnlyFiles = "safe" then
         Select Case LCase(File.FileExt)
          Case "txt","swf","doc","xls","gif","jpg","jpeg","pdf","bmp","css","htm","html","xls","zip","tiff","tif","mpg","mpeg","wav","mp3","csv","wmv","wma","png","mid","eps","js","cab","rm","ram","mp4","spp"
            File.Save DestinationPath
         	Case else:
            sThisError= "<table><tr><td bgcolor=red><Font Color=white><B>You cannot upload files with the "&File.FileExt&" extension.</b></Font></td></tr></table>"
            response.write " "
          end select
        else
            File.Save DestinationPath
        end if
      next

	 End If

  ElseIf Form.State > 10 then
	 Const fsSizeLimit = &HD
	 Select case Form.State
		  case fsSizeLimit: 
           sThisError= 	"<table><tr><td bgcolor=red><Font Color=white><B>Source form size (" & Form.TotalBytes & "B) exceeds form limit (" & Form.SizeLimit & "B)</b></Font></td></tr></table>"
		       response.write " "
      case else 
           sThisError=  "<table><tr><td bgcolor=red><Font Color=white><B>Some form error.</b></Font></td></tr></table>"
	         response.write " "
   end Select
  End If'Form.State = 0 then
  '{b}get an unique upload ID for this upload script and progress bar.
  Dim UploadID, PostURL
  UploadID = Form.NewUploadID

  'Send this ID as a UploadID QueryString parameter to this script.
  PostURL = Request.ServerVariables("SCRIPT_NAME") & "?UploadID=" & UploadID'{/b}
%>
