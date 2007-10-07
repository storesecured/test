<!--#include file="Global_Settings.asp"-->
<% if request.querystring("Filename") <> "" then %>
<SCRIPT LANGUAGE = "JavaScript">

	<!--
	function setParentResults(resArg,resVal) {
		window.parent.opener.setResults(resArg, resVal);
		window.parent.close();}
	//-->

</SCRIPT>
<body onLoad=setParentResults('<%= request.queryString("returnArg") %>','<%= request.querystring("Filename") %>');>
<%
else
%>
<html>
<head>
	<meta name="Pragma" content="no-cache">
	<META HTTP-EQUIV="pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Easystorecreator</title>
  <script language="JavaScript">
<!--
var submitcount = 0;
function checkFields(theform)
{
if (submitcount == 0) {
	//sumbit form
	submitcount ++;

	return true;
	}
else
	{
		alert("Please wait the upload is currently in progress.");
		return false;
	}
}
//-->
</script>
<%
Destination_Folder = fn_get_sites_folder(Store_Id,"Images")
sOnlyFiles = "safe"
 %>
<!--#include file="include/huge_upload.asp"-->
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<form method="POST" ENCTYPE='multipart/form-data' action=file_uploader.asp?ReturnArg=<%= request.querystring("ReturnArg") %> name=upload_image>
<% if SizeUsage > 100 then %>
	<font color=red>You are not allowed to upload anything because you are currently over your size limit.
<% else %>
        <input type="file" name="Image1" size="35"><BR>
        <input type="submit" name="Image" value=Upload size="35" onSubmit=""return checkFields();"">
        <BR>
        <font face=verdana size=2>Please only select Upload once.  It may take up to a minute for your image to be uploaded.
        <BR><BR>
        Do not upload images which contain special characters such as the ampersand, comma, space, colon, semicolon, brackets, parenthesis, etc.

<% end if %>
</form>
</body>
	</html>
	<SCRIPT>
//Open window with progress bar.
function ProgressBar(){
  var ProgressURL
  ProgressURL = 'progress.asp?UploadID=<%=UploadID%>'

  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=150')

  return true;
}
</SCRIPT>
    <% if sFilename<>"" then
        response.redirect "file_uploader.asp?ReturnArg="&request.querystring("ReturnArg")&"&FileName="&server.URLEncode(sFilename)
    end if
   
end if %>
