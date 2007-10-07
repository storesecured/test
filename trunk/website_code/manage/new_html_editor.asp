<!--#include file="Global_Settings.asp"-->
<!--#include file="include\sub.asp"-->
<%
 

sTextFieldName="popup_field"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>HTML Editor</title>
	<!-- STEP 1: include js files -->
	<script language=JavaScript src='editor/scripts/innovaeditor.js'></script>
	<script>
	//STEP 9: prepare submit FORM function
	function Save(resArg)
		{

    Form1.submit()
		//STEP 11: Here we move the edited content into the form input we've prepared before.
		window.parent.opener.setResults(resArg, oEdit<%= sTextFieldName %>.getHTMLBody());
		window.parent.close();

		}
	function LoadContent()
		{
		
    oEdit<%= sTextFieldName %>.putHTML(window.parent.opener.getResults('<%= request.queryString("returnArg") %>'))
    }
	</script>
</head>

<!-- STEP 4: After this page is fully loaded (means that the the file
to be edited is ready in the hidden IFRAME)
then we have to copy the content from the IFRAME into the editor. 
We do this on event body onload, by calling a function named LoadContent() -->
<body style="font:10pt verdana,arial,sans-serif"
topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 marginheight=0 marginwidth=0>
<form method="post" name="SaveForm" ID=Form1>
<%

sStartValue=encodeHTML(sStartValue)

response.write "<textarea rows='12' name='"&sTextFieldName&"' cols='83' id='"&sTextFieldName&"'>"&sStartValue&"</textarea>"
%>
<script>
			//document.getElementById("popup_field").value=window.parent.opener.getResults('Store_Header');
			document.getElementById("<%= sTextFieldName %>").value=window.parent.opener.getResults('<%= request.queryString("returnArg") %>');
			var oEdit<%= sTextFieldName %> = new InnovaEditor("oEdit<%= sTextFieldName %>");
			oEdit<%= sTextFieldName %>.width="100%";//You can also use %, for example: oEdit<%= sTextFieldName %>.width="100%"
			oEdit<%= sTextFieldName %>.height="92%";

			oEdit<%= sTextFieldName %>.features=["Search","Cut","Copy","Paste","PasteWord","|","Undo","Redo","|",
					  "ForeColor","BackColor","|","Hyperlink","InternalLink","Bookmark","|",
					  "Image","Flash","Media","|","Table","Guidelines","Absolute","|","Numbering","Bullets","|",
					  "Characters","Line","BRK",
					  "StyleAndFormatting","|","TextFormatting","ParagraphFormatting","ListFormatting","BoxFormatting","CssTest",
                                          "FontName","FontSize","CustomTag","|",
					  "Bold","Italic","Underline","|","Strikethrough","Subscript","Superscript","|",
					  "Indent","Outdent","|","JustifyLeft","JustifyCenter","JustifyRight","JustifyFull","|","RemoveFormat","ClearAll","HTMLSource"];

			oEdit<%= sTextFieldName %>.arrStyle = [["BODY",false,"","background:#FFFFFF;font-family:Verdana,Arial,Helvetica;font-size:x-small;"],
						[".ScreenText",true,"Screen Text","font-family:Tahoma;"],
						[".ImportantWords",true,"Important Words","font-weight:bold;"],
						[".Highlight",true,"Highlight","font-family:Arial;color:red;"]];

			oEdit<%= sTextFieldName %>.cmdAssetManager = "modalDialogShow('/Editor/assetmanager/assetmanager.asp',640,465)"; //Command to open the Asset Manager add-on.
  			oEdit<%= sTextFieldName %>.cmdInternalLink = "window.open('editor_internal.asp','mywindow','resizable=0,width=450,height=400')"; //Command to open your custom link lookup page.

			<% if sCustomTags="" then
				sCustomTags="[""None Available"",""""]"
			end if %>
			oEdit<%= sTextFieldName %>.arrCustomTag=[<%= sCustomTags %>];//Define custom tag selection
			oEdit<%= sTextFieldName %>.customColors=["#ff4500","#ffa500","#808000","#4682b4","#1e90ff","#9400d3","#ff1493","#a9a9a9"];//predefined custom colors
			oEdit<%= sTextFieldName %>.publishingPath="<%= Site_Name %>";
			oEdit<%= sTextFieldName %>.REPLACE("<%= sTextFieldName %>");
               //LoadContent();

		    //if (oUtil.setEdit) oUtil.setEdit("oEdit<%= sTextFieldName %>");


		</script>

	<CENTER><BR><INPUT type="button" value="Save Changes" onclick="Save('<%= request.queryString("returnArg") %>')" ID="Button1" NAME="Button1">

</body>
</html>

