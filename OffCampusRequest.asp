<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="include/StudentCacheAll.inc" -->
<!--#include file="include/AuthenticatePlugin.inc" -->
<!--#include file="include/EmailPlugin.inc" -->
<%
'THIS IS ON PRODUCTION
PageURL = "OffCampusRequest.asp?" & Request.QueryString


Call PageAuthenticate()
Call UpdateSessionActivity()


	Dim strRequest, permitType, attrStart, attrEnd, oTerm, fileOnly, appealText 'parameter passed in, type of permit requested, attribute start date, attribute end date, term object
	Dim FunctionKey, oFunction, oStudent
	Const permitStatus = "Pending"
	Const permitTypeAttrID = 8
	Const permitSemAttrID = 7
	Const permitAppAttrID = 6
	Const cancelReasonID = 2
	Const cancelReasonAttributeID = 4026
	Const cancelDateAttributeID = -23
	Const cancelReason = "Applying for an off-campus permit"
	FunctionKey = Request.QueryString("Function")
	
	Set oFunction = New CFunction
	oFunction.Initialize FunctionKey
	Set oTerm = New CTerm
	oTerm.Initialize oFunction.TermKey
	Set oStudent = New CStudent
	oStudent.Initialize PatronID
	
	
	Dim sFrom, sTo, sCCList, sBCCList, sSubject, sBody, sAttach
	sFrom = oStudent.Email
	sTo = "leeke1@sfasu.edu"
	
	sCCList = ""
	sBCCList = ""
	sSubject = "Exemption Request " & oTerm.Name
	
	
	strRequest = Request.QueryString("group1") '-- if something was passed to the file querystring
	fileOnly = Request.QueryString("fileOnly") '-- will be yes if this is an attempt to retrieve file only and not submit request
	appealText = Request.QueryString("appeal") '-- default is blank, can test again <> ""
	If strRequest <> "" Then 'get path of the file, note they are relative paths
	
		Select Case strRequest
			Case "Commuter-NacGrad"
				permitType = "Commuter-NacGrad"
				strRequest = "nacgrad.pdf"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "Commuter-NonNacGrad"
				permitType = "Commuter-NonNacGrad"
				strRequest = "nonnacgrad.pdf"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "Commuter-Relative"
				permitType = "Commuter-Relative"
				strRequest = "immediate_relative.pdf"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "Commuter-Property"
				permitType = "Commuter-Property"
				strRequest = "ownproperty.pdf"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "CustodyOfChild"
				permitType = "CustodyOfChild"
				strRequest = "custodyofchild.pdf"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "Married"
				permitType = "Married"
				strRequest = "married.pdf"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "Internet"
				permitType = "Internet"
				strRequest = ""
				attrStart = oTerm.StartDate
				attrEnd = oTerm.EndDate
			Case "Eight"
				permitType = "8"
				strRequest = "Eight_hour.pdf"
				attrStart = oTerm.StartDate
				attrEnd = oTerm.EndDate
			Case "8"
				permitType = "8"
				strRequest = "Eight_hour.pdf"
				attrStart = oTerm.StartDate
				attrEnd = oTerm.EndDate
			Case "Exemption"
				permitType = "Exemption"
				attrStart = oTerm.StartDate
				attrEnd = 0
			Case "Summer"
				permitType = "Summer-Agreement"
				attrSTart = oTerm.StartDate
				attrEnd = oTerm.EndDate
			
		End Select
		'
		
		If fileOnly = "" Then 'only write attribute and add to waiting list on initial submission, not subsequent file requests
			Dim oHMS, bSuccess, bSuccess2, vFailedRows 
			Set oHMS = GetClass("HMSDBSrv.PatronWrite")
			
			bSuccess = oHMS.InsertAttributeValue(OdysseyToken(), oStudent.Key, permitTypeAttrID, attrStart, attrEnd, permitType, vFailedRows)
			bSuccess2 = oHMS.InsertAttributeValue(OdysseyToken(), oStudent.Key, permitAppAttrID, attrStart, attrEnd, permitStatus, vFailedRows)
			bSuccess2 = oHMS.InsertAttributeValue(OdysseyToken(), oStudent.Key, permitSemAttrID, attrStart, attrEnd, oTerm.Name, vFailedRows) 
			
			If bSuccess = False Or Err.Number <> 0 Then
			  sError = sError & "Error Recording Permit Type Attribute: " & Err.Description & " " & GetFailedError(vFailedRows) & "||"
			End If
			Err.Clear


			'August 2013 - new request that off campus permit check to
			'see if the student has an application for the term and cancel
			'it.  This will cause the same result as if the student submitted
			'a cancellation and chose off-campus permit as the reason.
			Dim oApplications, rsApps, bAppUpdate, oTerms, rsTerms, appDefKey
			Dim vAttrFailedRows
			
			Set oApplications = GetClass("HMSDBSrv.ApplicationRead")
			Set rsApps = oApplications.GetPatronApplication(OdysseyToken(), , , , , , oFunction.TermKey)

			If rsApps.EOF = True Then
				'student doesn't have any applications for the term, that's ok
				'we just need a noop here or reverse the logic
			ElseIf rsApps("ApplicationType_Key") <> 6 Then
				appDefKey = rsApps("ApplicationDefinition_Key")
				Set oApplications = GetClass("HMSDBSrv.ApplicationWrite")
				
				'TODO if the application is already cancelled, don't update the cancelDate
				bAppUpdate = oApplications.WebUpdatePatronApplication(OdysseyToken(), vFailedRows, , oStudent.Key, appDefKey, , , cancelReasonID, , , Now())
				
				bSuccess = oHMS.InsertAttributeValue(OdysseyToken(), oStudent.Key, cancelReasonAttributeID, oTerm.StartDate, oTerm.EndDate, cancelReason, vAttrFailedRows)
				bSuccess2= oHMS.InsertAttributeValue(OdysseyToken(), oStudent.Key, cancelDateAttributeID, oTerm.StartDate, oTerm.EndDate, Now(), vAttrFailedRows)
				'rsApps("CancelledDate") = Now()
				'bAppUpdate = oApplications.UpdatePatronApplication(OdysseyToken(), oStudent.Key, rsApps, vFailedRows)
			End If
			
			If bAppUpdate = False Then
				Response.Redirect "Error.asp?Msg=" & Err.Description & " " & GetFailedError(vFailedRows)
			End If
					
			'A request was made to insert student address into the email, 
			'so we will pull that here before email is sent
			'address type 2 is the permanent address
			Dim rsAddress, sAddress
			Dim oAddressRead
			Set oAddressRead = CreateObject("HMSDBSrv.AddressRead")
			Set rsAddress = oAddressRead.GetAddresses2(OdysseyToken(), oStudent.Key, 2)
			sAddress = rsAddress("Street1") & vbCrLf
			sAddress = sAddress & rsAddress("City") & vbCrLf
			sAddress = sAddress & rsAddress("State") & vbCrLf
			sAddress = sAddress & rsAddress("ZIP") & vbCrLf
			
			
			'
			'Send email and insert note on exemption request
			'
				
				
			'do email stuff here
			If oStudent.Email <> "" Then
				sBody = "Student Name: " & oStudent.Name & vbCrLf & "Student ID: " & oStudent.ID & vbCrLf
				
				If appealText = "" Then
					sBody = sBody & "Submitted Off Campus Permit Request for reason: " & permitType & vbCrLf
				Else
					sBody = sBody & appealText
				End If
				
				If Len(sBody) > 0 And Len(sTo) > 0 Then
				  On Error Resume Next:
				  'we append the address only in the email
				  bSuccess = SendEmail(sFrom, sTo, sCCList, sBCCList, sSubject, sBody & sAddress)
				  If bSuccess = False Or Err.Number <> 0 Then
					sError = sError & "Error Sending Email Confirmation. Your request email confirmation could not be sent. Error details: " & Err.Description & "||"
					sError = sError & "However, your request was submitted successfully.||"
				  End If
				  'On Error Goto 0:
				End If

			Else: ' no student email
				sError = sError & "Error Sending Email Confirmation: Your request email confirmation could not be sent because you have no registered email address.||"
				sError = sError & "However, your request was submitted successfully.||"
			End If
			
			'send student email with form attached instead of download
			'this idea was abandoned, we still open popup with file for download
			'Dim attachFile
			'sBody = "You chose permit type " & permitType
			'sTo = "pestallworth@sfasu.edu"
			'sSubject = "Off Campus Permit Application"
			'attachFile = "D:\OdysseyWeb\HMSWEBStudent\test.txt"
			'bSuccess = SendEmail(sFrom, sTo, sCCList, sBCCList, sSubject, sBody)
			
			
			
			'
			' insert note into patron account
			'
			Dim vFailedRows2, rsNoteTypes, recCounter, noteKey
			Set oHMS = GetClass("HMSDBSrv.NoteRead")
			set rsNoteTypes = Server.CreateObject("ADODB.RecordSet")
			set rsNoteTypes = oHMS.GetNoteTypes(OdysseyToken())
			
			For recCounter = 1 to rsNoteTypes.RecordCount
				If rsNoteTypes.Fields.Item("NoteType_ID") = 10 Then
					noteKey = rsNoteTypes.Fields.Item("NoteType_Key")
				End If
				rsNoteTypes.MoveNext
			Next
			
			set oHMS = GetClass("HMSDBSrv.NoteWrite")

			'update 11/17/2010 to use sBody instead of appealText as the note body
			bSuccess = oHMS.InsertPatronNote(OdysseyToken(), oStudent.Key, noteKey, sBody, , , , vFailedRows2)
			
			If bSuccess = False Then
			  sError = sError & "Error Recording Note: " & Err.Description & " " & GetFailedError(vFailedRows2) & "||"
			End If
			' end of note section
			 

				' Log that student submitted request
				If Len(sSaveError) > 0 Then
				 Call StudentLog(oStudent.Key, "Page Error Student Request Details: " & sSaveError, Nothing)
				Else
					If Len(sSaveError) = 0 And Application("LogLevel") > 1 Then StudentLog oStudent.Key, "Log: Saved Off Campus Request, Term " & oTerm.Name & ",  Application Type " & oFunction.AppTypeName, Nothing
						Response.Redirect "Permit_Complete.asp?Function=" & FunctionKey & "&Permit=" & permitType
				End If
				Err.Clear
		End If 'end fileOnly block
		
		'
		'end of section from RequestDetail
		'

	End If 'end of strRequest <> ""
			
		
	If strRequest = "" Then 'this should be turned into an elseif above
	'the only way we will get here now is if someone types the link in directly, the string will be blank, we should simply catch and redirect home

		Response.redirect("https://odysseyweb.sfasu.edu/shsg/")
		Response.End
	End If

		
		
%>
</html>