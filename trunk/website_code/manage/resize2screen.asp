<%

on error goto 0

' This prefix will be added to every new image created by this script
' This prevents original images from overwriting
file_prefix = "_thmbnl"

' The following group of attributes works only if you defined
' to create html gallery
' This prefix will be added to every new image created by this script
' only in case if you are creating html gallery
thumb_prefix = "thumb_"
' Place attributes of html TABLE tag
html_table_attributes = "border=1 cellpadding=5 cellspacing=0"
' Place attributes of html TD tag
html_td_attributes = "align=middle valign=bottom"
html_horizontal = 0
html_vertical = 0
file_number = 0

' Set bWriteDate value to False if you don't want to print date when image was taken
bWriteDate = True
exif_xpos = 10
exif_ypos = 10
text_size = 21
text_color = "255,255,150"
text_outline_color = "0,0,0"
text_outline_width = 1

' Set to True, this value allows image auto rotation in case if exif data report
' what digicam had portrait orientation during image taking
' If digicam has no orientation sensor, the value does nothing
bCanRotate = True

all_colors = Array()


Set fs = CreateObject("Scripting.FileSystemObject")
Set g = CreateObject("shotgraph.image")

active_prefix = file_prefix

	source_dir = Folder_Path

	target_dir = source_dir
	' For wscript engine, hardcode maximal width and height
	maxwidth = sMaxWidth
	maxheight = sMaxHeight


' Add backslash symbol to target directory name, if required
if InStrRev(target_dir,"\") <> Len(target_dir) then
	target_dir = target_dir & "\"
end if

' This loop enumerates all files in the directory
cFiles = 0
sStart=request.querystring("Start")
if sStart="" then
   sStart=0
end if
sEnd=sStart+50
Set dir = fs.GetFolder(source_dir)
for each f in dir.Files
       if cint(cFiles)>=cint(sStart) and cint(cFiles)<=cint(sEnd) then

        ' Check if file has no our prefix
	if InStr(1,f.Name,active_prefix,1) < 1 then

                   ' Combine full path to source file
		   source_file = source_dir
		   ' Add backslash symbol to source file directory name, if required
		   if InStrRev(source_file,"\") <> Len(source_file) then
			source_file = source_file & "\"
		   end if
		   source_file = source_file & f.Name
		   ' Call the sub to resize file
                   source_file=replace(source_file,"\\","\")
                   MakeFile source_file,maxwidth,maxheight




	end if
	elseif cint(cFiles)>cint(sEnd) then
                    exit for
        end if
 cFiles = cFiles + 1
Next


if html_horizontal > 0 and html_vertical > 0 then
	CloseHtmlFile True

end if

'''''''''''''''''''''''''''''''''''''''
' END OF SCRIPT EXECUTION
'''''''''''''''''''''''''''''''''''''''


Sub MakeFile(filename,maxwidth,maxheight)

itype = g.GetFileDimensions(filename,width,height)
' Check if error occured
if itype <= 0 then Exit Sub



' Standard method to get new width and height
' to fit resized image into specified reclangle keeping aspect ratio
if width/maxwidth > height/maxheight then
	newwidth = maxwidth
	newheight = height * maxwidth \ width
else
	newheight = maxheight
	newwidth = width * maxheight \ height
end if

' Create primary imagespace
g.CreateImage newwidth,newheight,256

g.InitClipboard width,height
g.SelectClipboard True
if bRotate then
	' if image was rotated, get rotated image earlier prepared in g2
	g.ReadImage g2,palette,0,0
	Set g2 = Nothing
else
	g.ReadImage filename,palette,0,0
end if
' Resize image from secondary imagespace to primary one
g.Stretch 0,0,newwidth,newheight,0,0,width,height,"SRCCOPY","HALFTONE"
g.SelectClipboard False



' Combine new filepath using target directory, prefix and filename
pos = InStrRev(filename,"\")
g.JpegImage 70,0,target_dir & replace(Mid(filename,pos + 1),".",active_prefix&".")
if html_horizontal > 0 and html_vertical > 0 then
	AddFileToGallery target_dir,Mid(filename,pos + 1),newwidth,newheight
end if
End Sub


Sub PrintUsage()
text = "Resize to screen" & Chr(13) & Chr(10) &_
".VBS script for resizing digicam photos to specified size" &_
Chr(13) & Chr(10) &_
"Usage: cscript resize2screen.vbs <source directory> <target_directory> <size to feet images>" &_
" [<gallery page size>]" &_
Chr(13) & Chr(10) &_
"Examples: " &_
Chr(13) & Chr(10) &_
"cscript resize2screen.vbs c:\images ""c:\new images"" 1024x768" &_
Chr(13) & Chr(10) &_
"cscript resize2screen.vbs c:\images ""c:\new images"" 80x80 5x4"
Wscript.echo text
Wscript.Quit(1)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
' GetColor
' Returns color number from string value R,G,B
''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetColor(str)
For i=0 to UBound(all_colors)
	if all_colors(i) = str then GetColor = i: Exit Function
Next
a=Split(str,",")
i = UBound(all_colors) + 1
ReDim Preserve all_colors(i)
g.SetColor i,a(0),a(1),a(2)
all_colors(i) = str
GetColor = i
End Function

' This procedure writes outlined text on imagespace
Sub OutlineText(text, width, x, y, textColor, outlineColor)
g.SetBkMode "TRANSPARENT"
g.SetTextColor outlineColor
for i=1 to width
	g.TextOut x-i,y-i,text,True
	g.TextOut x+i,y+i,text,True
	g.TextOut x+i,y-i,text,True
	g.TextOut x-i,y+i,text,True
	g.TextOut x-i,y,text,True
	g.TextOut x+i,y,text,True
	g.TextOut x,y-i,text,True
	g.TextOut x,y+i,text,True
Next
g.SetTextColor textColor
g.TextOut x,y,text,True
End Sub

Sub AddFileToGallery(directory,filename,width,height)
filePosition = file_number mod (html_horizontal * html_vertical)

nGallery = file_number / (html_horizontal * html_vertical) + 1

if filePosition = 0 then
	CloseHtmlFile False
	Set html_file = fs.CreateTextFile(directory & "gallery" & nGallery & ".htm",True)
	html_file.WriteLine "<TABLE " & html_table_attributes & ">"
end if
if filePosition mod html_horizontal = 0 then
	if filePosition <> 0 then html_file.WriteLine "</TR>"
	html_file.WriteLine "<TR>"
end if

if filename <> "" then
html_file.WriteLine Chr(9) & "<TD " & html_td_attributes & ">" &_
"<A HREF=""" & filename & """ target=""_blank"">" &_
"<IMG SRC=""" & file_prefix & filename & """ width=" & width & " height=" & height & " border=1>" &_
"</A><BR>" & filename &_
"</TD>"
else
html_file.WriteLine Chr(9) & "<TD " & html_td_attributes & ">&nbsp;</TD>"
end if

file_number = file_number + 1
End Sub

Sub CloseHtmlFile(bFinalClose)
if file_number = 0 then Exit Sub

While file_number mod html_horizontal > 0
	AddFileToGallery "","",0,0
Wend

html_file.WriteLine "</TR>"
html_file.WriteLine "</TABLE>"

nGallery = (file_number-1) \ (html_horizontal * html_vertical) + 1
html_file.WriteLine "<BR><BR>"
html_file.WriteLine "<TABLE align=""center"" border=0 cellpadding=4 cellspacing=0>"
html_file.WriteLine "<TR>"

if nGallery > 1 then
html_file.WriteLine Chr(9) & "<TD><A HREF=""gallery" & (nGallery-1) & ".htm"">Previous</A></TD>"
else
html_file.WriteLine Chr(9) & "<TD></TD>"
end if

if bFinalClose then
html_file.WriteLine Chr(9) & "<TD></TD>"
else
html_file.WriteLine Chr(9) & "<TD><A HREF=""gallery" & (nGallery+1) & ".htm"">Next</A></TD>"
end if

html_file.WriteLine "</TR>"
html_file.WriteLine "</TABLE>"

html_file.Close
End Sub
%>
