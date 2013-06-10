Public Function Main()
	' Place VBScript code here
	Dim oPatron, oPatronList, oFactory, oCItems, oPatronNotes, sNote, bRetVal, sFilename, oFSO, oFile, sInline, sAlert, dAlertDate

	Set oFactory = HMS.PatronFactory

	sNote = InputBox("Enter note Text",,"Enter Note Text")
	If sNote = "" Then Exit Function
	
	sAlert = MsgBox("Set Note Alert?",4,"Alerted?")

	If sAlert = 6 Then
		dAlertDate = HMS.InputDateBox("Date to end alert",Now())
	End If

	sFilename = HMS.InputFilename("Choose file")
	If sFilename = "" Then Exit Function

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.OpenTextFile(sFilename)

	Do Until oFile.AtEndOfStream

		sInline = oFile.ReadLine
		Set oPatron = oFactory.GetByIDNumber(sInline)
		Set oPatronNotes = oPatron.Notes
		bRetVal = oPatronNotes.Add(7, sNote)
		If sAlert = 6 Then
			oPatronNotes.Alert = True
			oPatronNotes.AlertExpiration = dAlertDate
		End If
		oPatron.Save

	Loop
		
End Function
