Sub AttachmentPrint(Item As Outlook.MailItem)

On Error GoTo OError

#This script finds the system's Temp folders, saves any attachments, and runs the Print command for that file.

	Dim oFS As FileSystemObject
	Dim sTempFolder As String
	Set oFS = New FileSystemObject
	sTempFolder = oFS.GetSpecialFolder(TemporaryFolder)

	cTmpFld = sTempFolder & "\OETMP" & Format(Now, "yyyymmddhhmmss")
	MkDir (cTmpFld)
	
	#In the next few lines, you'll see an entry that says FileType
	#This line gets the last 4 characters of the file name, which we'll use later.

	Dim oAtt As Attachment
	For Each oAtt In Item.Attachments
		FileName = oAtt.FileName
		FileType = LCase$(right$(FileName, 4))
		FullFile = cTmpFld & "\" & FileName
		oAtt.SaveAsFile (FullFile)
		
	#We're using the FileType text. Note that it's the last 4 characters of the file name, 
	#which is why the next chunk has .xls and xlsx without the period and the period counts as the fourth character.
    #Insert any file extensions you want printed.

		Select Case FileType
		Case ".doc", "docx", ".xls", "xlsx", ".ppt", "pptx", ".pdf"
			Set objShell = CreateObject("Shell.Application")
			Set objFolder = objShell.NameSpace(0)
			Set objFolderItem = objFolder.ParseName(FullFile)
			objFolderItem.InvokeVerbEx ("print")
		End Select
	Next oAtt

	If Not oFS Is Nothing Then Set oFS = Nothing
	If Not objFolder Is Nothing Then Set objFolder = Nothing
	If Not objFolderItem Is Nothing Then Set objFolderItem = Nothing
	If Not objShell Is Nothing Then Set objShell = Nothing

	OError:
	If Err <> 0 Then
		MsgBox Err.Number & " - " & Err.Description
		Err.Clear
	End If
	
Exit Sub
End Sub

