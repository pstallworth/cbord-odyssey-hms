Public Function Main()
	' Place VBScript code here
	Dim oPatron, oPatronList, oFactory, oCItems, oPatronNotes, sNote, bRetVal, sFilename, oFSO, oFile, sInline, sAlert, dAlertDate
	Dim oTransCodes, retValues, oTranCode, arrText, count

	Set oTransCodes = HMS.TransactionCodes

	For i = 1 To oTransCodes.Count
		If oTransCodes.Code <> "Reset Hsg Cr" Then
			oTransCodes.MoveNext
		Else
			Exit For
		End If	
	Next

	'oTransCodes = "Reset Hsg Cr"

	Set oFactory = HMS.PatronFactory
	sFilename = HMS.InputFilename("Choose file")
	If sFilename = "" Then Exit Function


	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.OpenTextFile(sFilename)

	Do Until oFile.AtEndOfStream

		sInline = oFile.ReadLine
		arrText = Split(sInline, ",")
		Set oPatron = oFactory.GetByIDNumber(arrText(0))
		sCreditAmount = arrText(1)
		'MsgBox "Credit amount is: " & sCreditAmount
		retValues = oPatron.AddTransaction (sCreditAmount, oTransCodes.Code, , ,"FT Reset")
		
	Loop
		
End Function