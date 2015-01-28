<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="include/StudentCacheAll.inc" -->
<!--#include file="include/AuthenticatePlugin.inc" -->
<!--#include file="include/EmailPlugin.inc" -->
</head>
<body>
<%
On Error Resume Next
'Paul Stallworth, SFASU Room Swap Request function, Summer 2012
'This code allows for students to submit room swap requests via the 
'Odyssey WebStudent interface, and allows for the swapping of two students
'if a mutual request occurs.

'Some of the error handling could be better, right now it is a bit scattered
'and needs to be consolidated both in where it is located in the code and
'also in how the conditions are responded to.  The input validation stuff is 
'okay in my opinion, but other tests on if the assignments are valid for swap
'could be moved to the bottom with the rest of the response code
Call PageAuthenticate()
Call UpdateSessionActivity()

Const ATTRIB_SWAP_REQUEST = 3761

Dim oStudent, oFunction, oTerm
'All this preamble is required to setup the page and session info
'plus it gives you access to other variables like term dates, etc
'The last two lines initialize the student object for the currently logged in
'student submitting the request
'In this file, oStudent is the student on the page submitting the request and
'swapID and swapKey are for the ID of the student that oStudent has requested
'to swap with
	FunctionKey = Request.QueryString("Function")
	Set oFunction = New CFunction
	oFunction.Initialize FunctionKey
	Set oTerm = New CTerm
	oTerm.Initialize oFunction.TermKey
	Set oStudent = New CStudent
	oStudent.Initialize PatronID

'oPatronWrite, oPatronRead, oAttributeRead are all used for reading
'	and writing to students via the API

'swapID and swapKey are the ID number and key of the requested student
'	read in from the input box

'swapStudentSwapRequest is the swap request of the student associated
'	with swapID/swapKey

'swapEmail is email address associated with swapID

Dim oPatronWrite, oPatronRead, oAttributeRead, swapID, rs, vFailedRows
Dim bSuccess, bSuccess2, bSuccess4, swapStudentSwapRequest, swapKey, swapEmail
Dim swapFirstName, swapLastName
Dim sFacilityName, sFacilitySwapName, sError

'Saving a swap request and not a clear request
IF Request.Form("Save") = "1" AND Request.Form("ClearSwapRequest") <> "1" THEN
	swapID = Request.Form("swapid") 'swapID is ID entered in input box

	'client side checks for length, isnumeric, and others, also doing here
	'just to be redundant in case javascript not enabled on browser
	'server will notify if ID number not in database
	If Not IsNumeric(swapID) Then
         Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Error: Must enter numeric ID number.") 
    End If
	
	
	'Write the ID from input to student's attribute as their swap request
	Set oPatronWrite = GetClass("HMSDBSrv.PatronWrite")
	bSuccess = oPatronWrite.InsertAttributeValue(OdysseyToken(), oStudent.Key, ATTRIB_SWAP_REQUEST, oTerm.StartDate, oTerm.EndDate, swapID, vFailedRows)
	If bSuccess = False OR Err.Number <> 0 Then
		'this error msg can go away or be fixed
		sError = sError & "Error Recording Student ID for Swap: " & Err.Description & " " & GetFailedError(vFailedRows) & "||"
		Response.Redirect "Error.asp?Msg=" & sError & vbCrLf
	End If
	Err.Clear
	
	'At this point we have the PatronID but we need the Patron_Key for the next API call, so go fetch it
	Set oPatronRead = GetClass("HMSDBSrv.PatronRead")
	Set rs = oPatronRead.GetPatron(OdysseyToken(), ,swapID)
	If Err.Number <> 0 Then
		'Error occured, trap it
		Response.Redirect "Error.asp?Msg=Error looking up student: " & Err.Description
	End If
	On Error Goto 0
	If rs.RecordCount = 0 Then 'recordset is empty
		Response.Redirect "Error.asp?Msg=Error retrieving student information for swap request." & vbCrLf
	End If
	
	swapKey = rs("Patron_Key")
	swapEmail = rs("Email")
	swapFirstName = rs("FirstName")
	swapLastName = rs("LastName")
	
	'Use the StartDate + 1 to make sure it's inside the start of the term;
	'we start our term at 10AM so the +1 is done in case oTerm.StartDate ever
	'defaults to 00:00 as time
	'Get the corresponding swap request attribute for the swap student (swapID)
	'to see if they have mutually requested to swap
	Set oAttributeRead = GetClass("HMSDBSrv.AttributeRead")
	Set rs = oAttributeRead.BrowsePatronAttributesAtDate(OdysseyToken(), swapKey, oTerm.StartDate + 1)
	If rs.RecordCount = 0 Then 'recordset is empty, done
		'If this occurs no attributes were read, which would be
		'extremely odd indicating an error with the account.
		'At the very least it should return all the patron attributes
		'with no values.
	End If
	
	'Need to test what happens in next couple lines if this fails
	Do Until rs("Attribute_ID") = ATTRIB_SWAP_REQUEST 'attribute id for the Swap Request 
		rs.MoveNext
	Loop
	
	'swapStudentSwapRequest = 0 'initializing for future test
	swapStudentSwapRequest = rs("Value")


	Dim rsElements, lFacilityToKey, lElementKEy, lElementSwapKey, oContractRead, sAssetType, bCanSwap, lFacilityKey
	bCanSwap = True
	If oStudent.ID = swapStudentSwapRequest Then 'the patron's requested swap has also requested to swap with them
		
		'Verify both students are assigned for the term, and that the swap is allowed
		
		'Get the contract item element key for the term (will have to include state information to get preliminary only)
		'Could have many items in different states, but the only one we'll accept is a preliminary assignment
		'Also grabbing a couple other attributes from the record set to do some checks
		Set oContractRead = GetClass("HMSDBSrv.ContractRead")
		Set rsElements = oContractRead.GetContractElements(OdysseyToken(), oStudent.Key, oTerm.StartDate, oTerm.EndDate) 'change back to oTerm.StartDate
		
		IF Err.Number <> 0 Then
			Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: unable to retrieve patron assignment.")
		End IF
		Err.Clear
		
		lElementKey = 0
		lElementSwapKey = 0
		lFacilityToKey = 0
		Do Until rsElements.EOF
			If Not IsNull(rsElements("Facility_Key")) Then
				'Response.write "Facility_Key not null<br>" & vbCrLf
				If (rsElements("State_ID") = 1) Then 'This is the preliminary assignment
					lElementKey = rsElements("Element_Key") 'contract element of student logged in submitting request
					sAssetType = rsElements("AssetTypeName")
					sFacilityName = rsElements("Name") 'name of contract element
					lFacilityKey = rsElements("Facility_Key")
					Exit Do
				End If
			End If
			rsElements.MoveNext
		Loop
		
		IF lElementKey = 0 Then 'didn't find assignment key, break away here
			'If we get here I think something bad happened
			'non assigned students shouldn't be getting this far
			bCanSwap = False
			Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: No assignment found.")
		End If
		
		'first check if room is marked for special use
		Set rs = oAttributeRead.BrowseFacilityAttributesAtDate(OdysseyToken(), lFacilityKey, oTerm.StartDate)
		IF Err.Number <> 0 Then
			Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: Unable to retrieve facility information.")
		End If
		Err.Clear
		
		Dim sSpecialUse, sStudentFacilityType, sSwapStudentFacilityType
		sSpecialUse = "No"
				
		Do Until rs.EOF
			If Not IsNull(rs("FacilityAttribute_Key")) Then
				If (rs("Name") = "Special Use") Then
					sSpecialUse = rs("Value")
				ElseIf (rs("Name") = "Student Facility Type") Then
					sStudentFacilityType = rs("Value")
				End IF
			End IF
			rs.MoveNext
		Loop
		
		IF sSpecialUse = "Yes" Then
			Response.Redirect "Error.asp?Msg=Room not eligible for swap"
		ElseIf sAssetType = "CA" OR sAssetType = "LJ" OR sAssetType = "PC" THEN
			Response.Redirect "Error.asp?Msg=Room unavailable for swap (asset type)"
			'Error here
		END IF
	
		sAssetType = vbEmpty
		'Get all student contract items for the term
		'loop through and get the preliminary one (can only have one)
		'also grab a couple attributes of the assignment
		Set rsElements = oContractRead.GetContractElements(OdysseyToken(), swapKey, oTerm.StartDate, oTerm.EndDate) 'Stallworth change back to oTerm.StartDate
		
		IF Err.Number <> 0 Then
			Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: unable to retrieve patron assignment.")
		End IF
		Err.Clear
		
		Do Until rsElements.EOF
			If Not IsNull(rsElements("Facility_Key")) Then
				If (rsElements("State_ID") = 1) Then 'This is the preliminary assignment
					lFacilityToKey = rsElements("Facility_Key") 'facility where student is moving to
					lElementSwapKey = rsElements("Element_Key") 'contract item of other student that is swapping
					sAssetType = rsElements("AssetTypeName")
					sFacilitySwapName= rsElements("Name") 'name of contract element
					Exit Do
				End If
			End If
			rsElements.MoveNext
		Loop
		
		'This checks to see if the requested student is assigned,
		'if not we should finish saving the request but no swap
		'will be performed
		IF lElementSwapKey = 0 Then 'didn't find key, break away here
			bCanSwap = False
			Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: No assignment available for swap.")
		End If
		
		'first check if room is marked for special use
		Set rs = oAttributeRead.BrowseFacilityAttributesAtDate(OdysseyToken(), lFacilityToKey, oTerm.StartDate)
		IF Err.Number <> 0 Then
			Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: Unable to retrieve facility information.")
		End If
		Err.Clear
		
		sSpecialUse = "No"
		Do Until rs.EOF
			If Not IsNull(rs("FacilityAttribute_Key")) Then
				If (rs("Name") = "Special Use") Then
					sSpecialUse = rs("Value")
				ElseIf (rs("Name") = "Student Facility Type") Then
					sSwapStudentFacilityType = rs("Value")
					
				End IF
			End IF
			rs.MoveNext
		Loop
		
		IF sSwapStudentFacilityType <> sStudentFacilityType THEN
			Response.Redirect "Error.asp?Msg=Cannot swap with student in that facility."
		ElseIf sSpecialUse = "Yes" Then
			Response.Redirect "Error.asp?Msg=Room not eligible for swap"
		ElseIf sAssetType = "CA" OR sAssetType = "LJ" OR sAssetType = "PC" THEN
			Response.Redirect "Error.asp?Msg=Room unavailable for swap (asset type)"
			'Error here
		END IF
		
		'Do a final error check before attempting the switch
		'If sErrors = "" Then
		Dim oContractWrite, bSuccess3, vFailedRows2
		Set oContractWrite = GetClass("HMSDBSrv.ContractWrite")

		Const ATTRIB_SELECT_SPACE = 3111
		Dim sSelectSpace, bAttribSuccess, bAssignSuccess, sSwapSelectSpace
		Dim sAssignEnd, dAssignStart, dAssignEnd
		'sAssignEnd = CDate(#8/30/2012 9:55 AM#)
		'dAssignStart = CDate(#8/30/2012 10:05 AM#)
		'dAssignEnd = CDate(#12/15/2012 2:00 PM#)
		sAssignEnd = oTerm.EndDate
		dAssignStart = oTerm.StartDate
		dAssignEnd = oTerm.EndDate
		
		Set oAttributeRead = GetClass("HMSDBSrv.AttributeRead")
		'change the oTerm.StartDate to Now() or whatever is needed based 
		'on the time of year that the room swaps are occurring
		Set rs = oAttributeRead.BrowsePatronAttributesAtDate(OdysseyToken(), oStudent.Key, oTerm.StartDate)
		If rs.RecordCount = 0 Then 'recordset is empty, done
			'If this occurs no attributes were read, which would be
			'extremely odd indicating an error with the account.
			'At the very least it should return all the patron attributes
			'with no values.
		End If
		
		'Need to test what happens in next couple lines if this fails
		Do Until rs("Attribute_ID") = ATTRIB_SELECT_SPACE 'attribute id for the Swap Request 
			rs.MoveNext
		Loop
		
		sSelectSpace = rs("Value")
		
		Set rs = oAttributeRead.BrowsePatronAttributesAtDate(OdysseyToken(), swapKey, Now())
		If rs.RecordCount = 0 Then 'recordset is empty, done
			'If this occurs no attributes were read, which would be
			'extremely odd indicating an error with the account.
			'At the very least it should return all the patron attributes
			'with no values.
		End If
		
		'Need to test what happens in next couple lines if this fails
		Do Until rs("Attribute_ID") = ATTRIB_SELECT_SPACE 'attribute id for the Swap Request 
			rs.MoveNext
		Loop
		
		sSwapSelectSpace = rs("Value")

'========================================================================		
'When this was originally written, the room swap was occurring after the
'start of the semester and had to take into account ending the current
'assignment and then starting the new assignment after that.
'This does not need to be considered when the room swaps are occurring
'in a future semester, so only the last branch is needed where the rooms
'are swapped, and that is it.  You will have to comment in/out the other
'parts as needed depending upon the situation.

		'In the call below, the two ElementKeys specify not only the assignment but also the patron
		'lFacilityToKey is the facility (room) to move the first patron to
		'lElementKey is the assignment (contract item) to change/move for the first patron
		'lElementSwapKey is the contract item for the second patron to swap with the first
		If bCanSwap = True Then 'both assignments exist, we can swap
			'bSuccess3 = oContractWrite.ChangeAssignment3(OdysseyToken(), oTerm.StartDate ,lElementKey, lFacilityToKey, vFailedRows2, , , ,False , , lElementSwapKey) 
			'After swap, clear both student attributes so one student can't reverse the swap
			If sSelectSpace <> "Yes" AND sSwapSelectSpace <> "Yes" Then 'neither patron has done a swap
			'do the full set of updates and changes
				bAssignSuccess = oContractWrite.UpdateOneContractElement2(OdysseyToken(), lElementKey, vpFailedRows, , sAssignEnd)
				bAssignSuccess = oContractWrite.UpdateOneContractElement2(OdysseyToken(), lElementSwapKey, vpFailedRows, , sAssignEnd)
				bAssignSuccess = oContractWrite.AssignGroup(OdysseyToken(), lFacilityToKey, oStudent.Key, 1, dAssignStart, dAssignEnd, True, bUseRoommateLevel)
				bAssignSuccess = oContractWrite.AssignGroup(OdysseyToken(), lFacilityKey, swapKey, 1, dAssignStart, dAssignEnd, True, bUseRoommateLevel)
			Elseif sSelectSpace = "Yes" AND sSwapSelectSpace <> "Yes" Then 'current has done a swap
			'do update and assign for other patron, only assign for current patron
				bAssignSuccess = oContractWrite.UpdateOneContractElement2(OdysseyToken(), lElementSwapKey, vpFailedRows, , sAssignEnd)
				bAssignSuccess = oContractWrite.AssignGroup(OdysseyToken(), lFacilityKey, swapKey, 1, dAssignStart, dAssignEnd, True, bUseRoommateLevel)
				bAssignSuccess = oContractWrite.AssignGroup(OdysseyToken(), lFacilityToKey, oStudent.Key, 1, dAssignStart, dAssignEnd, True, bUseRoommateLevel)
			Elseif sSelectSpace <> "Yes" AND sSwapSelectSpace = "Yes" Then 'other has done a swap
			'do update and assign for current patron, only assign other patron
				bAssignSuccess = oContractWrite.UpdateOneContractElement2(OdysseyToken(), lElementKey, vpFailedRows, , sAssignEnd)
				bAssignSuccess = oContractWrite.AssignGroup(OdysseyToken(), lFacilityToKey, oStudent.Key, 1, dAssignStart, dAssignEnd, True, bUseRoommateLevel)
				bAssignSuccess = oContractWrite.AssignGroup(OdysseyToken(), lFacilityKey, swapKey, 1, dAssignStart, dAssignEnd, True, bUseRoommateLevel)
			Elseif sSelectSpace = "Yes" AND sSwapSelectSpace = "Yes" Then 'both have done a swap
			'do swap only
				bAssignSuccess = oContractWrite.ChangeAssignment3(OdysseyToken(), dAssignStart ,lElementKey, lFacilityToKey, vFailedRows2, , , ,False , , lElementSwapKey) 
			End If
			
			Dim  vAttrFailedRows
			Set oPatronWrite = GetClass("HMSDBSrv.PatronWrite")
			'write yes to both select space attributes now, also clear
			'both patrons swap request to prevent one of them from undoing
			'the change
			Dim bThrowAway
			If bAssignSuccess = True Then
				bSuccess3 = oPatronWrite.InsertAttributeValue(OdysseyToken(), oStudent.Key, ATTRIB_SELECT_SPACE, oTerm.StartDate, oTerm.EndDate, "Yes", vAttrFailedRows)
				bSuccess3 = oPatronWrite.InsertAttributeValue(OdysseyToken(), swapKey, ATTRIB_SELECT_SPACE, oTerm.StartDate, oTerm.EndDate, "Yes", vAttrFailedRows)
			Else
				Response.Redirect "Error.asp?Msg=" & Server.URLEncode("Internal Error: Not able to perform assignment swap.")
			End If
			Dim bClearSuccess
			Set oPatronWrite = GetClass("HMSDBSrv.PatronWrite")
			bClearSuccess = oPatronWrite.InsertAttributeValue(OdysseyToken(), oStudent.Key, ATTRIB_SWAP_REQUEST, oTerm.StartDate, oTerm.EndDate, "", vFailedRows3)
			bClearSuccess = oPatronWrite.InsertAttributeValue(OdysseyToken(), swapKey, ATTRIB_SWAP_REQUEST, oTerm.StartDate, oTerm.EndDate, "", vFailedRows3)
			
	
'==========================================================================================					
	Else 'no match found, or only one student has made swap request at this point
		'Swap information recorded
		
	End If

Elseif Request.Form("ClearSwapRequest") = "1" THEN
	'User has chosen to clear out their swap request, all we need to do is lookup and delete attribute value for current user
	Dim vFailedRows3
	Set oPatronWrite = GetClass("HMSDBSrv.PatronWrite")
	bSuccess4 = oPatronWrite.InsertAttributeValue(OdysseyToken(), oStudent.Key, ATTRIB_SWAP_REQUEST, oTerm.StartDate, oTerm.EndDate, "", vFailedRows3)
		
END IF 'outer most level

If bSuccess4 = True Then
	Response.Redirect "Default.asp?MsgSuccess=Swap request information has been cleared."
ElseIf bSuccess = True AND bSuccess3 = True Then
	'Writing of attributes and swapping of students both successful
	'Need to send email to both students and write notes

	Dim bEmailSuccess, sFrom, sTo, sCCList, sBCCList, sSubject, sBody, sAttach
	
	'Send email to oStudent making request 
	sFrom = "reslife@sfasu.edu"
	sTo = oStudent.Email
	sCCList = "reslife@sfasu.edu"
	sBCCList = ""
	sSubject = "Room Swap Request for " & oTerm.Name
	sError = ""
	
	If oStudent.Email <> "" Then
		sBody = "Student Name: " & oStudent.Name & vbCrLf & "Student ID: " & oStudent.ID & vbCrLf
		sBody = sBody & "You successfully swapped from " & sFacilityName & " to " & sFacilitySwapName & "." & vbCrLf & vbCrLf
		sBody = sBody & "Please check the facility rates on our website www.sfasu.edu/reslife/101.asp as changing rooms may affect your bill." & vbCrLf & vbCrLf
		sBody = sBody & "Thank you," & vbCrLf & "SFA Residence Life Department" & vbCrLf
		
		If Len(sBody) > 0 And Len(sTo) > 0 Then
		  bEmailSuccess = SendEmail(sFrom, sTo, sCCList, sBCCList, sSubject, sBody)
		  If bEmailSuccess = False Or Err.Number <> 0 Then
		  	'sError = sError & "Error Sending Email Confirmation. Your request email confirmation could not be sent. Error details: " & Err.Description & "||"
			Call StudentLog(oStudent.Key, "Page Error RoomSwapRequest, Term " & oTerm.Name & ", " & Err.Description, Nothing)
		  End If
		  'On Error Goto 0:
		End If

	Else: ' no student email
		sError = sError & " Error Sending Email Confirmation: Your request email confirmation could not be sent because no email address is listed in the system."
	End If
	
	Err.Clear
	
	'This gets the note key for use at the end, we are writing a General Note for swaps at this time
	'This just gets all note types in the system, loops until it finds the one with the matching ID, then
	'pulls internal (primary) key for it
	Dim oHMS, rsNoteTypes, recCounter, noteKey
	Set oHMS = GetClass("HMSDBSrv.NoteRead")
	set rsNoteTypes = Server.CreateObject("ADODB.RecordSet")
	set rsNoteTypes = oHMS.GetNoteTypes(OdysseyToken())

	For recCounter = 1 to rsNoteTypes.RecordCount
		If rsNoteTypes.Fields.Item("NoteType_ID") = 1 Then
			noteKey = rsNoteTypes.Fields.Item("NoteType_Key")
		End If
		rsNoteTypes.MoveNext
	Next
	
	Dim oPatronNote, noteBody, bSuccess5
	set oPatronNote = GetClass("HMSDBSrv.NoteWrite")
	noteBody = "Student swapped from " & sFacilityName & " to " & sFacilitySwapName & sError & ". ODYWEB"
	If sError <> "" Then
		noteBody = noteBody & sError
	End IF
	
	bSuccess5 = oPatronNote.InsertPatronNote(OdysseyToken(), oStudent.Key, noteKey, noteBody, , , , vFailedRows2)
	
	'=========================================================================
	'Now send email and insert note for swapStudent
	Dim bEmailSuccess2, sError2
	sFrom = "reslife@sfasu.edu"
	sTo = swapEmail
	sCCList = "reslife@sfasu.edu"
	sBCCList = ""
	sSubject = "Room Swap Request for " & oTerm.Name
	sError2 = ""
	If sTo <> "" Then
		sBody = "Student Name: " & swapFirstName & " " & swapLastName & vbCrLf & "Student ID: " & swapID & vbCrLf
		sBody = sBody & "You successfully swapped from " & sFacilitySwapName & " to " & sFacilityName & "." & vbCrLf & vbCrLf
		sBody = sBody & "Please check the facility rates on our website www.sfasu.edu/reslife/101.asp as changing rooms may affect your bill." & vbCrLf & vbCrLf
		sBody = sBody & "Thank you," & vbCrLf & "SFA Residence Life Department" & vbCrLf
		
		If Len(sBody) > 0 And Len(sTo) > 0 Then
		  bEmailSuccess2 = SendEmail(sFrom, sTo, sCCList, sBCCList, sSubject, sBody)
		  If bEmailSuccess2 = False Or Err.Number <> 0 Then
		  	'sError = sError & "Error Sending Email Confirmation. Your request email confirmation could not be sent. Error details: " & Err.Description & "||"
			Call StudentLog(swapKey, "Page Error RoomSwapRequest, Term " & oTerm.Name & ", " & Err.Description, Nothing)
		  End If
		  'On Error Goto 0:
		End If

	Else: ' no swap student email, don't display anything to current user
		sError2 = "Error Sending Email Confirmation: No Email address on file"
	End If
	
	'This gets the note key for use at the end, we are writing a General Note for swaps at this time
	'This just gets all note types in the system, loops until it finds the one with the matching ID, then
	'pulls internal (primary) key for it
	
	Set oHMS = GetClass("HMSDBSrv.NoteRead")
	set rsNoteTypes = Server.CreateObject("ADODB.RecordSet")
	set rsNoteTypes = oHMS.GetNoteTypes(OdysseyToken())

	For recCounter = 1 to rsNoteTypes.RecordCount
		If rsNoteTypes.Fields.Item("NoteType_ID") = 1 Then
			noteKey = rsNoteTypes.Fields.Item("NoteType_Key")
		End If
		rsNoteTypes.MoveNext
	Next
	
	
	noteBody = "Student swapped from " & sFacilitySwapName & " to " & sFacilityName & " " & sError2 & ". ODYWEB"
	set oPatronNote = GetClass("HMSDBSrv.NoteWrite")

	
	bSuccess5 = oPatronNote.InsertPatronNote(OdysseyToken(), swapKey, noteKey, noteBody, , , , vFailedRows2)
	'========================================================================
	Call StudentLog(oStudent.Key, "Room Swap Completed", Nothing)
	Call StudentLog(swapKey, "Room Swap Completed", Nothing)
	Response.Redirect "Default.asp?MsgSuccess=Successfully completed room swap. " & sError
Elseif bSuccess = True AND bSuccess3 = False Then
	'Wrote swap request to oStudent, but cannot make swap or swap failed
	'Most likely reasons are
	'a)swapStudent doesn't have assignment
	'b)swapStudent hasn't mutually requested a swap with oStudent
	'c)either student is in a room that does not allow swapping
	
	Response.Redirect "Default.asp?MsgSuccess=Swap request saved, no swap performed."
Else
	'Need to send email and write note for student submitting request
	Response.Redirect "Default.asp?MsgSuccess=Swap request information has been saved."
End If

%>
</body>
</html>
