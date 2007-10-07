<%
	'bigdownload.asp - download of files with size up to 2GB
	'Sample for ScriptUtils.ByteArray
	'c2003, http://www.pstruh.cz

	'SendFileByBlocks "F:\work\WebLib.dll", 65535

	Sub SendFileByBlocks(FileName, BlockSize)
		Dim FileSize, ByteCounter
		FileSize = CreateObject("scripting.filesystemobject").GetFile(FileName).Size

		'Switch off buffer.
		Response.Buffer = False

		'This is download
		Response.ContentType = "application/x-msdownload"

		'Set file name
		Response.AddHeader "Content-Disposition", _
		  "attachment; filename=""" & GetFileName(FileName) & """"

		'Set Content-Length (ASP doen not set it when Buffer = False)
		Response.AddHeader "Content-Length", FileSize
		'Response.CacheControl = "no-cache"

		Dim BA
		Set BA = CreateObject("ScriptUtils.ByteArray")

		'Loop through file contents.
		For ByteCounter = 1 To FileSize Step BlockSize
			'Do not write data when client is disconnected
			If Not Response.IsClientConnected() Then Exit For

			'Read block of data from a file
			BA.ReadFrom FileName, ByteCounter, BlockSize

			'Write the block to output.
			Response.BinaryWrite BA.ByteArray
		Next
		set BA = Nothing
		Set FileSize = nothing
		
	End Sub


	Function GetFileName(FullPath)
		Dim Pos, PosF
		PosF = 0
		For Pos = Len(FullPath) To 1 Step -1
			Select Case Mid(FullPath, Pos, 1)
				Case ":", "/", "\": PosF = Pos + 1: Pos = 0
			End Select
		Next
		If PosF = 0 Then PosF = 1
		GetFileName = Mid(FullPath, PosF)
	End Function
%>

