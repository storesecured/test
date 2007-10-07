<STYLE>body{background-color:buttonface;color:buttontext;}</STYLE>
<TITLE>Color Wheel</TITLE>
<SCRIPT LANGUAGE = "JavaScript">
	<!--
	selectedFName = 'Bgcolor';
	function displayNewColors() {}
	function setNewValue (cvalue) {
		if (selectedFName == 'Bgcolor') {
		  document.hex.Bgcolor.value = cvalue;}
		displayNewColors()}
	function setNewFromForm (cvalue) {
		if (selectedFName == 'Bgcolor') {
		  document.Color_Test.bg.value = document.hex.Bgcolor.value;}
		displayNewColors() }
	function setSelected (fname) {
		selectedFName = fname}

	function setParentResults(resArg) {
		self.opener.setResults(resArg, document.hexform.hex.value);
		window.close();}



//-->
</SCRIPT>

<SCRIPT language="vbScript">

' Author:	Lewis Edward Moten III
' Email: Lewis@Moten.com
' URL:		http://www.lewismoten.com/
' ICQ:		308364
' Created:	02/11/2002

' Description:
'	Scriptlet to allow users to choose
'	a color.  A color wheel is displayed
'	with an assortment of colors - most of
'	witch are web-safe.	Clicking on a color
'	causes it to be outlined and raise an event
'	that a color was chosen.  This object
'	mimics one of the color-picker forms found
'	in Microsoft Office 2000.

' Requirements:
'	Internet Explorer 5.5 or later
'	Possible need for DirectX.  Version unknown.

Dim lngY	' Top position of image
Dim lngX	' Left position of image
Dim lngIndex	' Image Index
Dim strColors	' String of all Hex color values

' Color Wheel
strColors = _
							"0033663366993366CC0033990000990000CC000066" & _
						"0066660066990099CC0066CC0033CC0000FF3333FF333399" & _
					"00808000999933CCCC00CCFF0099FF0066FF3366FF3333CC666699" & _
				"33996600CC9900FFCC00FFFF33CCFF3399FF6699FF6666FF6600FF6600CC" & _
			"33993300CC6600FF9966FFCC66FFFF66CCFF99CCFF9999FF9966FF9933FF9900FF" & _
		"00660000CC0000FF0066FF9999FFCCCCFFFFCCECFFCCCCFFCC99FFCC66FFCC00FF9900CC" & _
	"00330000800033CC3366FF6699FF99CCFFCCFFFFFFFFCCFFFF99FFFF66FFFF00FFCC00CC660066" & _
		"33660000990066FF3399FF66CCFF99FFFFCCFFCCCCFF99CCFF66CCFF33CCCC0099800080" & _
			"33330066990099FF33CCFF66FFFF99FFCC99FF9999FF6699FF3399CC3399990099" & _
				"66663399CC00CCFF33FFFF66FFCC66FF9966FF7C80FF0066D60093993366" & _
					"808000CCCC00FFFF00FFCC00FF9933FF6600FF5050CC0066660033" & _
						"996633CC9900FF9900CC6600FF3300FF0000CC0000990033" & _
							"663300996600CC3300993300990000800000A50021"

' Grays
strColors = strColors & _
						"F8F8F8DDDDDDB2B2B28080805F5F5F3333331C1C1C080808" & _
							"EAEAEAC0C0C09696967777774D4D4D292929111111"

' Black & White 
strColors = strColors & "FFFFFF000000"

' Write hex cubes
For lngIndex = 0 to 143
	
	' If hex is small
	If lngIndex < 142 Then
		document.write "<IMG src=""images/SmallHex.gif"" style=""filter:progid:DXImageTransform.Microsoft.Light();position:absolute;width:14px;height:15px;display:none;"">"
	Else ' big hex
		document.write "<IMG src=""images/BigHex.gif"" id=""Hex" & lngIndex & """ style=""filter:progid:DXImageTransform.Microsoft.Light();position:absolute;width:28px;height:31px;display:none;"">"
	End If

	' update X/Y positions for each row
	Call SetXY()
	
	' Move into proper position, set color, etc.
	Call SetupHex()
next

' Write out ring (to outline chosen color)
document.write "<IMG src=""images/SmallHexRing.gif"" style=""position:absolute;top:320px;left:320px;display:none;"" id=""HexRing"">"

' Write out text "Colors:"
document.write "<SPAN style=""font-family:Arial;font-size:10pt;position:absolute;top:8px;left:7px;"">Colors:</SPAN>"
document.write "<br><br><br><br><br><br><br><br><br><br><br><br><form name=hexform>"
document.write "<input type=text name=hex id=hex value=FFFFFF size=10>"
Sub Document_onClick()

	Dim llngTop	' top position of clicked hex
	Dim llngLeft	' left position of clicked hex
	Dim lobjHEX	' reference to hex clicked
	Dim lblnBig	' user clicked big hex?

	' If user didn't click on a hex cube ...	GET OUT!
	If Not window.event.srcElement.tagName = "IMG" Then Exit Sub
	If window.event.srcElement.id = "HexRing" Then Exit Sub
	
	' Hide hex-ring
	HexRing.style.display = "none"
	
	' get reference to hex cube clicked
	Set lobjHEX = window.event.srcElement
	
	' parse top and left positions (sometimes have text "px" within them.)
	llngTop = Replace(lobjHEX.style.top, "px", "")
	llngLeft = Replace(lobjHEX.style.left, "px", "")
	
	' did user click on a big hex?
	lblnBig = Right(lobjHEX.src, 11) = "/BigHex.gif"
	If lblnBig Then
		' Load ring
		HexRing.src = "images/BigHexRing.gif"
		' Resize
		HexRing.style.width = 35
		HexRing.style.height = 39
		' Position around hex color clicked
		HexRing.style.top = llngTop - 5
		HexRing.style.left = llngLeft - 4
	Else ' clicked on small hex
		' Load ring
		HexRing.src = "images/SmallHexRing.gif"
		' Resize
		HexRing.style.width = 21
		HexRing.style.height = 23
		' Position around hex color clicked
		HexRing.style.top = llngTop - 4
		HexRing.style.left = llngLeft - 3
	End If
	
	' make that ring come back
	HexRing.style.display = "inline"
 
	
	Dim TheForm
	Set TheForm = Document.hexForm
	TheForm.hex.Value = lobjHEX.alt 
  
	' Remove short term memory
	Set lobjHEX = Nothing
	
	' huh?
	
End Sub

Sub Window_onLoad()
	' Use windows colors to give appearance of belonging to O.S.
	document.body.style.backgroundColor = "buttonface"
	document.body.style.color = "buttontext"
End Sub

Sub SetupHex()

	Dim llngRed	' red luminance
	Dim llngGreen	' green luminance
	Dim llngBlue	' blue luminance
	Dim lobjImg	' Reference to image address
	Dim lstrHex	' RGB hex color value
	
	' fetch hex color
	lstrHex = Mid(strColors, (lngIndex * 6) + 1, 6)

	' parse decimal values
	llngRed 	= CInt("&h" & Left(lstrHex, 2))
	llngGreen	= CInt("&h" & Mid(lstrHex, 3, 2))
	llngBlue = CInt("&h" & Right(lstrHex, 2))
	
	' Get reference address of hex element ... (Shorten path reference)
	Set lobjImg = document.images(lngIndex)
	
	' Magical filters ...
	Call lobjImg.filters.item(0).addAmbient(llngRed, llngGreen, llngBlue, 100)
	
	' Move It!
	lobjImg.style.left = lngX
	lobjImg.style.top = lngY
	
	lobjImg.alt = lstrHex ' save hex in alternate text for reference (vb - tag)
	
	' Ok - its finnished.  Let's display it.
	lobjImg.style.display = "inline"
	
	' Take some garbage out of the computers short term memory
	Set lobjImg = Nothing
	
	' Prepare for the next hex
	lngX = lngX + 14
	
End Sub

Sub SetXY()
	' This is where we change the X/Y position
	' of the next hex color cube depending on the
	' index of the image.
	Select Case lngIndex
		' Color Wheel
		Case 0:lngX=61:lngY=26
		Case 7:lngX=54:lngY=38
		Case 15:lngX=47:lngY=50
		Case 24:lngX=40:lngY=62
		Case 34:lngX=33:lngY=74
		Case 45:lngX=26:lngY=86
		Case 57:lngX=19:lngY=98
		Case 70:lngX=26:lngY=110
		Case 82:lngX=33:lngY=122
		Case 93:lngX=40:lngY=134
		Case 103:lngX=47:lngY=146
		Case 112:lngX=54:lngY=158
		Case 120:lngX=61:lngY=170
		' Grays
		Case 127:lngX=54:lngY=198
		Case 135:lngX=61:lngY=210
		' White
		Case 142:lngX=19:lngY=194
		' Black
		Case 143:lngX=173:lngY=194
	End Select
End Sub
</SCRIPT>

<input value="set" type="button" onClick="JavaScript:setParentResults('<%= request.queryString("returnArg") %>');">

</form>
