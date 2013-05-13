<script language="VBScript" runat="server">

	Const cancelReason_ID = 4026
	Const sValidReason = "Not Attending SFA"

Public Function ApplyDecision(ByVal PatronKey, ByVal BCF, ByVal AssignmentDecision, _ 
				ByVal Decision,	byVal DepositDecision, ByVal startDate, _
				ByVal endDate, ByVal inventorySheet, ByRef vFailed) ' need to add ByRef vFailed
				
	' need attributes deposit process, processed, insert note
	' should also consider creation of BCF attribute
	Const DepositProcess_ID = 9937
	Const Processed_ID = 6936
	Const NoteType_ID = 11
	Const NoteType_Key = 31


	Dim oHMS, rsPatronAttr, bSuccess, bSuccess2, sNote

	Set oHMS = GetClass("HMSDBSrv.PatronWrite")
	bSuccess = True
	If DepositDecision <> "" Then
		bSuccess = oHMS.InsertAttributeValue(GetUserSession("StaffToken"), _
					PatronKey, DepositProcess_ID, startDate, endDate, _
					Left(DepositDecision, Len(DepositDecision) - 1), _
					vFailed)
	End If


	If bSuccess = False Then 
		ApplyDecision = False
	End If

	bSuccess2 = oHMS.InsertAttributeValue(GetUserSession("StaffToken"), _
					PatronKey, Processed_ID, startDate, endDate, _
					"Yes", vFailed)
	
	If Not (bSuccess AND bSuccess2) Then
		Err.Raise vbObjectError + 9999, "Apply Decision", "InsertAttributeValue() failed"
	End If
	
	If inventorySheet = True Then
		sNote = "Received inventory sheet. "
	End If
	
	sNote = sNote & "Cancellation Decision: " & Decision & " Deposit: " & DepositDecision & _
			" Assignment: " & AssignmentDecision & " BCF: " & BCF & " User: " & GetUserSession("StaffName")
			
	Set oHMS = GetClass("HMSDBSrv.NoteWrite")
	bSuccess = oHMS.InsertPatronNote(GetUserSession("StaffToken"), _
					PatronKey, NoteType_Key, sNote,,,, vFailed)
	
	ApplyDecision = bSuccess
	
End Function

' may need to look at adding an isenrolled() test at the bottom
' and just like hasApp, if it's true hold the deposit in case of loophole
' students
' 5/7/13: per discussion with D'nese, we need to look at the cancel reason
' and if the student is 21/60, and not process if not eligible to live off
' campus and reason is not "not attending"
' Also need to look at only returning cancellation / renewal cancellations

Sub DetermineDeposit(ByVal lPatronKey, ByVal lTermKey, _
								ByVal dCancelled, byVal TermIndicator, _
								ByVal dSubmittedDate, ByVal atDate, _
								ByRef Decision, ByRef DepositDecision _
								)

	Const sValidReason = "Not Attending SFA"
	Const cancelReason_ID = 4026
	Dim pDate, cType, cReason

	'moving these two up here to avoid multiple calls'
	cType = CurrentResidentCancellation(lPatronKey, lTermKey)
	cReason = GetCancelReason(lPatronKey, lTermKey, atDate)

	If IsSuspended(lPatronKey, atDate) = True Then
		sDecision = "Academic Suspension. §17C3"
		sDepositDecision = "Refund." ' Contract Sec. 17C3
	Elseif IsDisciplineSuspension(lPatronKey, atDate) = True Then
		sDecision = "Discipline suspension.  §17C1"
		sDepositDecision = "Forfeit." ' Contract Sec. 17C1
	Elseif IsRejected(lPatronKey, atDate) = True Then
		sDecision = "Not eligible for admission. §18"
		sDepositDecision = "Refund." ' Contract Sec 18
	ElseIf Not IsNull(cType) AND _
				StrComp(cReason, sValidreason, 1) <> 0 AND _
				(HasSixtyHours(lPatronKey, atDate) = False AND _
				IsTwentyOne(lPatronKey, lTermKey) = False) Then

		'Need to come full stop if either of these next two occur
		'as it means paperwork / data entry hasn't been done correctly
		Assert Not IsNull(cReason), "Cancel reason cannot be null or missing"
		Assert cReason <> "", "Cancel reason cannot be empty"

		'either 60/21 or applying for OC permit'
		If getOCPermitStatus(lPatronKey, atDate, , "") = "Pending" Then
			sDepositDecision = "Do not process."
			sDecision = "Student has pending OC permit."
		ElseIf getOCPermitStatus(lPatronKey, atDate, , "") = "Completed" Then
			sDepositDecision = "Refund."
			sDecision = "Student has OC permit."
		Else
			sDepositDecision = "Hold for review."
			sDecision = "Student does not have 60 hours, will not be 21 by semester."

		End If

	Elseif DateDiff("d", dCancelled, GetDeadline(lTermKey)) >= 0 Then
		If TermIndicator = "Future" Then ' 17A
			sDecision = "Cancelled on or before deadline. §17A"
			sDepositDecision = "Refund."
			If getOCPermitStatus(lPatronKey, atDate, , "") = "Pending" Then
				sDepositDecision = "Do not process." ' Not specified in contract
				sDecision = "Student has pending OC permit."
			End If
		Elseif TermIndicator = "Current" Then
			sDepositDecision = "Refund."
			If getOCPermitStatus(lPatronKey, atDate, , pDate) = "Completed" Then
				If DateDiff("d", pDate, GetClassesStart(lTermKey)) >= 0 Then
					'student cancelled before deadline and received permit before classes started
					sDepositDecision = "Refund." 'Contract Sec 17B2
					sDecision = "Student has OC permit."
				Else
					'student either cancelled after deadline or received permit after deadline
					sDepositDecision = "Forfeit."
					sDecision = "Student received OC permit after first day of classes."
				End If
			Elseif isGraduating(lTermKey, lPatronKey) = True Then
				sDepositDecision = "Refund." ' Contract Sec 17B6
				sDecision = "Student is graduating. §17B6"
			Else
				sDepositDecision = "Refund."
				sDecision = "Student cancelled before deadline..."
			End If
		End If
	Else
		sDepositDecision = "Forfeit." ' Contract Sec 17A
		If TermIndicator = "Future" Then
			If DateDiff("d", dCancelled, GetDeadline(lTermKey)) < 0 AND _
				DateDiff("d", dCancelled, atDate) >= 0 AND _
				DateDiff("d", dSubmittedDate, GetDeadline(lTermKey)) < 0 Then
	
				sDepositDecision = "Refund."
				sDecision = "Student cancelled before term, submitted application after deadline."
			Else
				sDepositDecision = "Forfeit."
				sDecision = "Student cancelled after deadline."
			End if
		Elseif TermIndicator = "Current" Then
			If DateDiff("d", dCancelled, GetDeadline(lTermKey)) < 0 AND _
				DateDiff("d", dCancelled, atDate) >= 0 AND _
				DateDiff("d", dSubmittedDate, GetDeadline(lTermKey)) < 0 Then
	
				sDepositDecision = "Refund."
				sDecision = "Student cancelled before term, submitted application after deadline."
			Else
				sDepositDecision = "Forfeit."
				sDecision = "Student cancelled after deadline."
			End If
		End If
	End If
End Sub

Public Function DetermineBreakContractFee(ByVal PatronKey, ByVal TermKey, ByVal TermIndicator, ByVal atDate, ByVal endDate)

	Const MaxAssignLengthBeforeCBF = 1
	Dim pDate
'	If IsSuspended(PatronKey, atDate) = True Then
'		DetermineBreakContractFee = "No, academic suspension." ' Contract Sec. 17C3
	If IsDisciplineSuspension(PatronKey, endDate) = True Then
		DetermineBreakContractFee = "No, discipline suspension. §17C1" ' Contract Sec. 17C1
	Elseif IsRejected(PatronKey, atDate) = True Then
		DetermineBreakContractFee = "No, not eligible for admission. §18" ' Contract Sec. 18, doesn't mention CBF
	Else
		If TermIndicator = "Future" Then
			DetermineBreakContractFee = "No §17A" ' Contract Section 17A
		Else
			If getOCPermitStatus(PatronKey, atDate, , pDate) = "Completed" Then
				If DateDiff("d", GetDeadline(TermKey), pDate) >= 0 AND _
					DateDiff("d", pDate, GetClassesStart(TermKey)) >= 0 Then
					'student received permit before classes start and cancelled before deadline
					DetermineBreakContractFee = "No, has OC permit.  §17B2"
				Else
					'student either cancelled after deadline or received permit after deadline
				DetermineBreakContractFee = "Yes." ' Contract Section 17B2
				End If
			Else
				If IsAssigned(PatronKey, atDate, endDate, "") = True Then
					Dim dActualStart, dActualEnd
					GetContractActualDates PatronKey, atDate, endDate, dActualStart, dActualEnd
					
					' If they remove their belongings on day facility opens,
					' then there is no BCF, otherwise there is.
					If dActualStart <> "" AND dActualEnd <> "" Then ' 
						If DateDiff("d", dActualStart, dActualEnd) < MaxAssignLengthBeforeCBF AND _ 
							DateDiff("d", dActualStart, atDate) <= MaxAssignLengthBeforeCBF Then
							DetermineBreakContractFee = "No, checked out on opening day. §17B4"
						Else
							DetermineBreakContractFee = "Yes, checked out during semester."
						End If
					Elseif dActualStart <> "" Then
						DetermineBreakContractFee = "Yes, student not yet checked out."
					Else
						DetermineBreakContractFee = "No, never checked in."
					End If
				Else
					DetermineBreakContractFee = "No, not assigned."
				End If
			End If
		End If
	End If
End Function

'atDate should be term start
Sub ProcessAssignment(ByVal PatronKey, ByVal TermKey, ByVal TermIndicator, ByVal startDate, ByVal atDate, ByRef AssignmentMessage)

	Dim contractState, actualStart, actualEnd

	If TermIndicator = "Future" Then
		If IsAssigned(PatronKey, startDate, startDate + 1, contractState) = True Then
			AssignmentMessage = "Cancel assignment, void charges."
		Else
			AssignmentMessage = "No assignment, verify no charges."
		End If
	Elseif TermIndicator = "Current" Then
		If IsAssigned(PatronKey, startDate, atDate, contractState) = True Then
			GetContractActualDates PatronKey, startDate, atDate, actualStart, actualEnd
			If contractState = "Completed" AND actualStart <> "" AND actualEnd <> "" Then 
				AssignmentMessage = "Contract already complete, "
			Else
				AssignmentMessage = "Complete contract, "
			End If

			If IsSuspended(PatronKey, startDate) = True Then
				AssignmentMessage = AssignmentMessage & "student on academic suspension, "
			End If
			
			If IsDisciplineSuspension(PatronKey, atDate) = True Then
				AssignmentMessage = AssignmentMessage & "discipline suspension, keep charges on account."
			Elseif getOCPermitStatus(PatronKey, startDate, , "") = "Completed" Then
				AssignmentMessage = AssignmentMessage & "has OC permit, prorate charges."
			Elseif getOCPermitStatus(PatronKey, startDate, , "") = "Pending" Then
				AssignmentMessage = "Do not process, student has pending off-campus permit, awaiting decision."
			Elseif IsTwentyOne(PatronKey, TermKey) = True Then
				AssignmentMessage = AssignmentMessage & "is 21 yrs old, prorate charges."
			Elseif HasSixtyHours(PatronKey, startDate) = True Then
				AssignmentMessage = AssignmentMessage & "has 60 earned hours, prorate charges."
			Elseif IsEnrolled(PatronKey, startDate) = False Then
				AssignmentMessage = AssignmentMessage & "is not enrolled, prorate charges."
			Else
				AssignmentMessage = "Don't change contract, don't change charges, student still enrolled. §17B5"
			End If

		Elseif IsAssigned(Patronkey, startDate, atDate, contractState) = 1 Then
			AssignmentMessage = "¡Multiple active assignments for term!"
		Else
			AssignmentMessage = "Student not assigned for term. Verify no charges on account."
		End If
	End If

End Sub

 
Public Function getOCPermitStatus(ByVal lPatronKey, ByVal atDate, ByVal endDate, ByRef permitDate)

	Dim oPatronAttr, rsPatronAttr
	Set oPatronAttr = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oPatronAttr.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
'	Set rsPatronAttr = oPatronAttr.GetPatronAttributeValues(GetUserSession("StaffToken"), lPatronKey, atDate)
'	rsPatronAttr.Filter = "Name='Permit Date'"
	rsPatronAttr.Filter = "Name='Permit App'"

	permitDate = ""
	
	If rsPatronAttr.EOF = True Then
		getOCPermitStatus = False
		Exit Function
	End If
	If rsPatronAttr("Value").Value = "C" OR rsPatronAttr("Value").Value = "Completed" Then
		rsPatronAttr.Filter = "Name='Permit Date'"
		If rsPatronAttr.EOF = False Then
			permitDate = rsPatronAttr("Value").Value
		End If
		getOCPermitStatus = "Completed"
	Elseif rsPatronAttr("Value").Value = "P" or rsPatronAttr("Value").Value = "Pending" Then
		
		getOCPermitStatus = "Pending"
	Else
		getOCPermitStatus = False
	End If

End Function

' TODO Stub
Public Function isGraduating(ByVal lTermKey, ByVal lPatronKey)

	isGraduating = False
	
End Function



' this function takes a term key and gets the next logical
' term key. e.g. if term is fall 2012, next term is spring 2013
Public Function GetNextTerm(ByVal lTermKey)

	Dim oTerms, rsTerms, sNextTermName, sFullNextTerm, sNextTermYear
	Dim dDeadline
	Set oTerms = GetClass("HMSDBSrv.TermRead")
	Set rsTerms = oTerms.GetTerms(GetUserSession("StaffToken"), lTermKey)
	
	rsTerms.Filter = "Term_Key=" & lTermKey
		
	Assert Not rsTerms.EOF, "Error looking up next term."
	Dim a
	a = Split(rsTerms("Description").Value)
	
	If a(0) = "Fall" Then 
		sNextTermName = "Spring"
		dDeadline = rsTerms("CancellationDeadline").Value
		sNextTermYear = CInt(a(1)) + 1
	ElseIf a(0) = "Spring" Then 
		sNextTermName = "Fall"
		dDeadline = rsTerms("CancellationDeadline").Value
		sNextTermYear = a(1)
	ElseIf a(0) = "Summer" Then
		sNextTermName = "Fall"
		dDeadline = rsTerms("CancellationDeadline").Value
		sNextTermYear = a(2)
		'sNextTermYear = a(1) & " " & a(2)
	End If
	
	sFullNextTerm = sNextTermName & " " & sNextTermYear
	rsTerms.Filter = "Description='" & sFullNextTerm & "'"
	If rsTerms.EOF <> True Then
		GetNextTerm = rsTerms("Term_Key").Value
	Else
		GetNextTerm = 0
	End If

End Function

' Deadline here means the cancellation deadline associated with
' the term.  The deadline must be defined in the HMS UI Setup module
' for the term.
Public Function GetDeadline(ByVal lTermKey)
	Dim oTerms, rsTerms
	Set oTerms = GetClass("HMSDBSrv.TermRead")
	Set rsTerms = oTerms.GetTerms(GetUserSession("StaffToken"), lTermKey)
	
	rsTerms.Filter = "Term_Key=" & lTermKey
		
	If rsTerms.EOF = True Then
		Response.redirect "Error.asp?Error=Must select a term for this to work."
	End If
	
	Assert Not IsNull(rsTerms("CancellationDeadline").Value), "Cancellation deadline is not defined!"
	
	GetDeadline = rsTerms("CancellationDeadline").Value

End Function

' GetContractActualDates
' we get the list of student contracts for the dates request, if no contracts,
' then we exit.  Otherwise we retrieve the actual start and end dates.  In
' the case of multiple active contracts, we report that instead to notify
' the user.
Sub GetContractActualDates(ByVal PatronKey, atDate, DisplayDate, ByRef ActualStart, ByRef ActualEnd)

	Dim oContractRead, rsElements, contractCount
	Set oContractRead = GetClass("HMSDBSrv.ContractRead")
	Set rsElements = oContractRead.GetContractElements(GetUserSession("StaffToken"), lPatronKey, atDate, DisplayDate)

	ActualStart = ""
	ActualEnd = ""
	contractCount = 0
	If rsElements.EOF = False Then 
		Do Until rsElements.EOF
			If Not IsNull(rsElements("Facility_Key")) Then
				If (rsElements("State_ID") = 1 Or rsElements("State_ID") = 2) Then
					contractCount = contractCount + 1
					ActualStart = rsElements("ActualUseStart").Value
					ActualEnd = rsElements("ActualUseEnd").Value
				End If
			End If
			rsElements.MoveNext
		Loop
	Else
		Exit Sub
	End If

	If contractCount = 1 Then
		Exit Sub
	Elseif contractCount > 1 Then 'Hmm, need to report this
		ActualStart = "Multiple active assignments"
		ActualEnd = "Multiple active assignments"
		Exit Sub
	End If
	
	rsElements.MoveFirst
	
	If rsElements.EOF = False Then 
		Do Until rsElements.EOF
			If Not IsNull(rsElements("Facility_Key")) Then
				If rsElements("State_ID") = 4 OR rsElements("State_ID") = 6 Then
					ActualStart = rsElements("ActualUseStart").Value
					ActualEnd = rsElements("ActualUseEnd").Value
					Exit Do
				End If
			End If
			rsElements.MoveNext
		Loop
	End If
	
End Sub

Public Function IsAssigned(ByVal PatronKey, ByVal atDate, ByVal endDate, ByRef contractState) 'missing a parameter here that is global

	Dim oElements, rsElements, contractCount
	Set oElements = GetClass("HMSDBSrv.ContractRead")
	Set rsElements = oElements.GetContractElements(GetUserSession("StaffToken"), PatronKey, atDate, endDate)

	' we want to look for active contracts first, if we find just one, then we're done
	' if we find more than one active contract, need to report that
	' if we find no active contracts but completed contracts, that is ok too
	
	contractCount = 0
	If rsElements.EOF = TRUE Then 
		IsAssigned = False
	Else
		Do Until rsElements.EOF
			If Not IsNull(rsElements("Facility_Key")) Then
					If (rsElements("State_ID") = 1 Or rsElements("State_ID") = 2) Then
						contractState = "Active"
						contractCount = contractCount + 1
						'IsAssigned = True
						'Exit Do
					End If
			End If
			rsElements.MoveNext
		Loop
		
		If contractCount = 1 Then
			IsAssigned = True
		Elseif contractCount > 1 Then 'Hmm, need to report this
			IsAssigned = 1
		End If
		
		rsElements.MoveFirst
		
		Do Until rsElements.EOF
			If Not IsNull(rsElements("Facility_Key")) Then
					If (rsElements("State_ID") = 4 OR rsElements("State_ID") = 6) Then
						contractState = "Completed"
						IsAssigned = True
						Exit Do
					End If
			End If
			rsElements.MoveNext
		Loop
	End If
	
	If contractState <> "Active" AND contractState <> "Completed" Then
		IsAssigned = False
	End If
	
End Function

Public Function IsSuspended(ByVal PatronKey, ByVal atDate)

	Dim oAttrRead, rsPatronAttr
	
	Set oAttrRead = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oAttrRead.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
	
	rsPatronAttr.Filter = "Name='Academic Standing'"
	If rsPatronAttr.EOF = False Then
		If rsPatronAttr("Value").Value = "Suspension" OR _
			rsPatronAttr("Value").Value = "Continue Suspension" Then 
			IsSuspended = True
		Else
			IsSuspended = False
		End If
	End if
	
End Function


Public Function IsDisciplineSuspension(ByVal PatronKey, ByVal atDate)

	Dim oAttrRead, rsPatronAttr
	
	Set oAttrRead = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oAttrRead.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
	
	rsPatronAttr.Filter = "Name='Discipline Suspension'"
	If rsPatronAttr.EOF = False Then
		If rsPatronAttr("Value").Value = "Yes" Then
			IsDisciplineSuspension = True
		Else
			IsDisciplineSuspension = False
		End If
	End if

End Function

Public Function IsRejected(ByVal PatronKey, ByVal atDate)

	Dim oAttrRead, rsPatronAttr
	
	Set oAttrRead = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oAttrRead.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
	
	rsPatronAttr.Filter = "Name='Admit Status'"
	If rsPatronAttr.EOF = False Then
		If rsPatronAttr("Value").Value = "RJ" OR _
			rsPatronAttr("Value").Value = "NE" OR _ 
			rsPatronAttr("value").Value = "AP" Then 
			IsRejected = True
		Else
			IsRejected = False
		End If
	End if

End Function

Public Function IsEnrolled(ByVal PatronKey, ByVal atDate)

	Dim oAttrRead, rsPatronAttr
	
	Set oAttrRead = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oAttrRead.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
	
	rsPatronAttr.Filter = "Name='Current Hours'"
	If rsPatronAttr.EOF = False Then
		If rsPatronAttr("Value").Value <> "" AND _
			rsPatronAttr("Value").Value > 0 Then
			IsEnrolled = True
		Else
			IsEnrolled = False
		End If
	End if
	
End Function

' University policy requires students under 21 to live on campus
Public Function IsTwentyOne(ByVal PatronKey, ByVal TermKey)

	Dim oHMS, rsTerms, censusDate, dAge, dAgeNow, rsPatron
	Set oHMS = GetClass("HMSDBSrv.TermRead")
	Set rsTerms = oHMS.GetTerms(GetUserSession("StaffToken"), TermKey)
	
	Set oHMS = GetClass("HMSDBSrv.PatronRead")
	Set rsPatron = oHMS.GetPatron(GetUserSession("StaffToken"), PatronKey)
	
	rsTerms.Filter = "Term_Key=" & TermKey
		
	' TODO This redirect needs to go away	
	If rsTerms.EOF = True Then
		Response.redirect "Error.asp?Error=Must select a term for this to work."
	End If
	
	censusDate = rsTerms("CensusDate").Value
	
	dAge = DateDiff("yyyy",rsPatron("Birthdate").Value,censusDate)
	
	If (Month(censusDate) > Month(rsPatron("Birthdate").value)) Then
		dAge = dAge - 1
	Elseif (Month(censusDate) = Month(rsPatron("Birthdate").value)) AND _
			Day(censusDate) > Day(rsPatron("Birthdate").value) Then
		dAge = dAge - 1
	End If
		
	If (dAge >= 21) Then 
		IsTwentyOne = True
	Else
		IsTwentyOne = False
	End If
	
End Function

' University policy requires students with < 60 hours to live on campus
Public Function HasSixtyHours(ByVal PatronKey, ByVal atDate)

	Dim oAttrRead, rsPatronAttr
	
	Set oAttrRead = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oAttrRead.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
	
	rsPatronAttr.Filter = "Name='Earned Hours'"
	If rsPatronAttr.EOF = False Then
		If rsPatronAttr("Value").Value <> "" AND _
			rsPatronAttr("Value").Value >= 60 Then 'University policy
			HasSixtyHours = True
		Else
			HasSixtyHours = False
		End If
	End If
	
End Function

' Looks for cancellations from current students, null otherwise
Public Function CurrentResidentCancellation(ByVal PatronKey, ByVal TermKey)

	Dim oHMS, rsPatronApplications
	Set oHMS = GetClass("HMSDBSrv.ApplicationRead")
	Set rsPatronApplications = oHMS.GetPatronApplication(GetUserSession("StaffToken"), ,PatronKey,,,,TermKey)
	
	rsPatronApplications.Filter = "(Term_Key="& TermKey & " AND " & _
			"SubmittedDate<>Null AND CancelledDate=Null AND ApplicationType_Key=6)"  & _
			" OR (Term_Key="& TermKey & " AND CancelledDate<>Null AND ApplicationType_Key=8)"
	
	If rsPatronApplications.EOF = True Then
		CurrentResidentCancellation = Null
	Else
		CurrentResidentCancellation = rsPatronApplications("ApplicationType_Key")
	End If
	
End Function

Public Function GetCancelReason(ByVal PatronKey, ByVal TermKey, ByVal atDate)

	Dim oHMS, oHMS2, rsPatronApplications, rsPatronAttr, cancelReason
	Set oHMS = GetClass("HMSDBSrv.ApplicationRead")
	Set rsPatronApplications = oHMS.GetPatronApplication(GetUserSession("StaffToken"), ,PatronKey,,,,TermKey)
	Set oHMS2 = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oHMS2.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), PatronKey, atDate)

	rsPatronApplications.Filter = "(Term_Key="& TermKey & " AND " & _
			"SubmittedDate<>Null AND CancelledDate=Null AND ApplicationType_Key=6)"  & _
			" OR (Term_Key="& TermKey & " AND CancelledDate<>Null AND ApplicationType_Key=8)"

	'if the cancellation is a returning cancellation, the cancel reason will
	'be in the attribute, not with the application
	If rsPatronApplications.EOF = True Then
		'no cancellation for types passed in
		'...how did we even get here?'
	Else
		GetCancelReason = rsPatronApplications("ApplCancelReasonName").Value
		If rsPatronApplications("ApplCancelReasonName").Value = "" OR _
			 IsNull(rsPatronApplications("ApplCancelReasonName").Value) Then
			'verify data correctly entered'
			If rsPatronApplications("ApplicationType_Key").Value = 8 Then
				'Force user to correct student record'
				Err.Raise vbObjectError + 9990, "Renewal Cancellation", "Incorrectly filled out renewal cancellation"
			Else
				'otherwise check the attribute'
				rsPatronAttr.Filter = "Attribute_ID=" & cancelReason_ID
				If IsNull(rsPatronAttr("Value").Value) OR rsPatronAttr("Value").Value = "" Then
					Err.Raise vbObjectError + 9991, "Returning Cancellation", "Missing cancel reason attribute"
				End If
				GetCancelReason = rsPatronAttr("Value").Value
			End If
		End If
	End If
End Function

Public Function GetCancelledDate(ByVal PatronKey, ByVal TermKey)

	Dim oHMS, rsPatronApplications
	Set oHMS = GetClass("HMSDBSrv.ApplicationRead")
	Set rsPatronApplications = oHMS.GetPatronApplication(GetUserSession("StaffToken"), ,PatronKey,,,,TermKey)
	
	rsPatronApplications.Filter = "(Term_Key="& TermKey & " AND CancelledDate<>Null" & _
			" AND ApplicationType_Key=1) OR (Term_Key="& TermKey & " AND " & _
			"SubmittedDate<>Null AND CancelledDate=Null AND ApplicationType_Key=6)"  & _
			" OR (Term_Key="& TermKey & " AND CancelledDate<>Null AND ApplicationType_Key=8)"

	If rsPatronApplications.EOF = True Then
		'Student doesn't have cancellation, so we need to break out here
		'if possible to avoid any more calculations
		GetCancelledDate = Null
	Else
		If rsPatronApplications("ApplicationType_Key") = 6 Then
			GetCancelledDate = rsPatronApplications("SubmittedDate").Value
		Else
			GetCancelledDate = rsPatronApplications("CancelledDate").Value
		End If
	End If
	
End Function

Public Function GetSubmittedDate(ByVal PatronKey, ByVal TermKey)

	Dim oPatronAttr, rsPatronApplications, txt
	Set oPatronAttr = GetClass("HMSDBSrv.ApplicationRead")
	Set rsPatronApplications = oPatronAttr.GetPatronApplication(GetUserSession("StaffToken"), ,PatronKey,,,,TermKey)


	rsPatronApplications.Filter = "(Term_Key="& TermKey & " AND " & "CancelledDate=Null" & _
			" AND ApplicationType_Key=1) OR (Term_Key="& TermKey & " AND " & _
			"CancelledDate=Null AND ApplicationType_Key=8)"  

	If rsPatronApplications.EOF = True Then
		'Student doesn't have cancellation, so we need to break out here
		'if possible to avoid any more calculations
		GetSubmittedDate = Null
	Else
		GetSubmittedDate = rsPatronApplications("SubmittedDate").Value
	End If
		
	
End Function

' This is similar to GetDeadline, it pulls the Classes Start Date from
' the term information in the HMS UI Setup module.  It must be set
' or an invalid result will be returned.
Public Function GetClassesStart(ByVal lTermKey)
	Dim oTerms, rsTerms
	Set oTerms = GetClass("HMSDBSrv.TermRead")
	Set rsTerms = oTerms.GetTerms(GetUserSession("StaffToken"), lTermKey)
	
	rsTerms.Filter = "Term_Key=" & lTermKey
	
	' This is not good, doing the redirect from here.
	If rsTerms.EOF = True Then
		Response.redirect "Error.asp?Error=Must select a term for this to work."
	End If
	
	Assert Not IsNull(rsTerms("ClassesStartDate").Value), "Classes Start Date is not defined!"
	
	GetClassesStart = rsTerms("ClassesStartDate").Value
End Function

Public Function getProcessedStatus(ByVal PatronKey, ByVal atDate)

	Dim oHMS, rsPatronAttr
	
	Set oHMS = GetClass("HMSDBSrv.AttributeRead")
	Set rsPatronAttr = oHMS.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, atDate)
	
	rsPatronAttr.Filter = "Name='Cancellation Processed'"
	If rsPatronAttr.EOF = False Then
		If rsPatronAttr("Value").Value <> "" AND _
			rsPatronAttr("Value").Value = "Yes" Then 'University policy
			getProcessedStatus = True
		Else
			getProcessedStatus = False
		End If
	End If
End Function

Public Function hasFutureApp(ByVal PatronKey, ByVal TermKey)

	Dim oHMS, rsApps, rsTerm, termName
	
	Set oHMS = GetClass("HMSDBSrv.ApplicationRead")
	Set rsApps = oHMS.GetPatronApplication(GetUserSession("StaffToken"), ,PatronKey)
	Set oHMS = GetClass("HMSDBSrv.TermRead")
	Set rsTerm = oHMS.GetTerms(GetUserSession("StaffToken"))
	
	rsTerm.Filter = "Term_Key='" & TermKey & "'"
	termName = rsTerm("Name")
	
	Do Until rsApps.EOF
		If CLng(rsApps("TermName").Value) > CLng(termName) Then
			If Not (IsNull(rsApps("SubmittedDate"))) AND _
					IsNull(rsApps("CancelledDate")) AND _
					rsApps("ApplicationType_Key") <> 6 Then
				hasFutureApp = True
				Exit Function
			End If
		End If
		rsApps.MoveNext
	Loop
	
	hasFutureApp = False
	
End Function

Sub Assert( boolExpr, strOnFail )
        If Not boolExpr Then
				Err.Raise vbObjectError + 9999, "PatronCustomForm1.asp", strOnFail
        End If
End Sub
</script>