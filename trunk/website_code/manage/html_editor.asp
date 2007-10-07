<!--#include virtual="common/global_settings.asp"-->
<!--#include file="include\sub.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>HTML Editor</title>
	<!-- STEP 1: include js files -->
	<script language="JavaScript" src="include/yusasp_ace.js"></script>
	<script language="JavaScript" src="include/yusasp_color.js"></script>
	<script>
	//STEP 9: prepare submit FORM function
	function Save(resArg)
		{
		//STEP 10: Before getting the edited content, the display mode of the editor 
		//must not in HTML view.
		if(obj1.displayMode == "HTML")
			{
			alert("Please uncheck HTML view")
			return ;
			}
    Form1.submit()
		//STEP 11: Here we move the edited content into the form input we've prepared before.
		window.parent.opener.setResults(resArg, obj1.getContentBody() );
		window.parent.close();

		}
	function LoadContent()
		{
		
    obj1.putContent(window.parent.opener.getResults('<%= request.queryString("returnArg") %>'))
    }
	</script>
</head>

<!-- STEP 4: After this page is fully loaded (means that the the file 
to be edited is ready in the hidden IFRAME) 
then we have to copy the content from the IFRAME into the editor. 
We do this on event body onload, by calling a function named LoadContent() -->
<body onload="LoadContent()" style="font:10pt verdana,arial,sans-serif" 
topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 marginheight=0 marginwidth=0>

<!-- STEP 6: Prepare a FORM to submit the edited content back to the server -->
<form method="post" name="SaveForm" ID=Form1>

<!-- STEP 7: Since we edit content within editor that is not part of the FORM,
we need to move the edited content into a FORM input. 
Here we prepare a hidden FORM input. -->
<input type="hidden" name="txtContent"  value="" ID=txtContent>
		
	<!-- STEP 2: Attach Editor (+ enable the asset function) -->
	<script>
	var obj1 = new ACEditor("obj1")
	obj1.width = "100%"
	obj1.height = "95%"
	obj1.isFullHTML = false //edit full HTML (not just BODY)
	obj1.ImagePageURL = "/editor_image.asp" //specify Image library management page
	obj1.AssetPageURL = "/editor_Asset.asp" //specify Asset library management page
	obj1.InternalLinkPageURL = "/editor_internal.asp" //specify Asset library management page
    obj1.InternalLinkPageHeight = "400"
    obj1.useZoom = true
    obj1.usePrint = false
    obj1.useStyle = false
    obj1.usePageProperties = false
	obj1.useInternalLink = true
    obj1.base = "<%= Site_Name2 %>" //where the users see the content (where we publish the edited content)
    obj1.baseEditor = "http://manage.easystorecreator.com/" //location of the editor
    obj1.storeid = "<%= Store_Id %>"

    
    obj1.RUN()
	</script>

	<INPUT type="button" value="S A V E" onclick="Save('<%= request.queryString("returnArg") %>')" ID="Button1" NAME="Button1">
  <font size=1>The HTML Editor requires Windows with IE5.0 or later. Some features are designed for IE5.5 or later</font>

</form>


</body>
</html>

