<%@ Language=VBScript %>
<% 
  Option Explicit

  Response.AddHeader "Pragma", "no-store"
  Response.CacheControl = "no-store"
  Response.Expires = -1500
 %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="include/form.inc"-->
<!--#include file="include/Table.inc" -->
<!--#include file="include/Patrons.inc" -->
<!--#include file="include/StaffCacheAll.inc" -->
<!--#include file="include/PageCodePatronCustomForm1.inc" -->

<!--#include file="include/CancellationForm.inc" -->
<%
	' Author: Paul Stallworth
	' Custom Form for processing Cancellations
	' I'm coding this very conservatively, explicitly declaring
	' variable storage to hold all my decision values at first.
	' Future revisions I may remove the storage and calculate
	' decisions based upon the recordsets themselves but I want
	' to ensure I get all business logic accounted for first,
	' in lieu of how l33t the code might be.
	
	
  ' PatronGeneral - displays general information for currently loaded patron (or no info if no patron selected).
  ' Expects these parameters:
  '   - UserSession("PatronKey") - currently loaded patron
  '   - QueryString("PatronKey")/QueryString("Key") - currently loaded patron
  '   - QueryString("Msg") - optional success message to user
  '   - QueryString("Error") - optional error message to user
  '   - QueryString("cmd") - optional, can be either "new" for new Patron or "save" to save core fields info
  ' For form submit:
  '   - Form("cmd") - "save"

  ValidateUser "PatronCustomForm1.asp"

	Server.ScriptTimeout = 420
  Dim oPage, oFunction
  Set oPage = New CPage
  oPage.InitializeByName "PatronCustomForm1"
  Set oFunction = oPage.PageFunction

  Call PageCodeHeaderPatronCustomForm1(oPage.Key)

  ' verify user has access to patrons
  If GetUserSession("UR_Read_PS") <> 1 Then
    Response.Redirect "Error.asp?Error=" & Server.URLEncode("Access Denied to " & GetPatronLabel() & ". Please see your administrator.")
  End If

  ' If PatronKey QueryString specified, load the specified patron if not already loaded
	Dim lPatronKey, sAssignmentMessage, sCmd, sErrors, sPatronName, sPatronID, bNoteAlert, sPatronTitle, sAltID
	Dim dStart, dEnd, lTermKey, dDisplayDate, dtNow
	Dim oHMS, rsPatronAttributes, rsPatronApplications, rsPatronApplications2, rsPatron, rsElements, bAssigned
	Dim dCancelledDate, dSubmittedDate, bSuspended, bRejected, bTwentyOne, dBirthdate
	Dim bSixtyHours, bEnrolled, sDepositDecision, sBCF, dAge, dCancelled, rsTerms
	Dim sNextTermName, sNextTermYear, sFullNextTerm, lNextTermKey, bGraduating, bHasOCPermit
	Dim sDecision, sHoldDeposit, sTermRelative, bHasFutureApp, dDeadline, dAssignedDate
	Dim dActualStart
  
  lTermKey = 0
  sErrors = Request.QueryString("Error")
  sAssignmentMessage = Request.QueryString("Msg")
  sCmd = LCase(Request.QueryString("cmd"))
  If Request.QueryString("TermKey") <> "" Then
     lTermKey = CLng(Request.QueryString("TermKey"))
  ElseIf GetUserSession("TermKey") <> "" Then 
     lTermKey = CLng(GetUserSession("TermKey"))
  End If
  lPatronKey = CLng(GetUserSession("PatronKey"))
  sPatronID = Request.QueryString("id")
  sPatronName = ""
  dtNow = Now()

  GetDisplayDates Request.QueryString("TermKey"), dStart, dEnd
 
	If GetUserSession("PatronKey") = "" Then
		Response.Redirect "PatronGeneral.asp?Error=" & Server.URLEncode("Please select a " & GetPatronLabel() & " first.")
	End If

  ' If a patron was looked up and sent to this page, load patron
  If Request.QueryString("PatronKey") <> "" Then
    LoadPatron Request.QueryString("PatronKey"), Nothing, sPatronName, sPatronID, sAltID, bNoteAlert, False
    lPatronkey = CLng(Request.QueryString("PatronKey"))
  Else:
    lPatronkey = CLng(GetUserSession("PatronKey"))
  End If
  sPatronTitle = GetPatronTitle()
  sPatronID = GetUserSession("PatronID")

  If lTermKey <> 0 Then
     If dStart >= dtNow Then
        dDisplayDate = DateAdd("n", 1439, DateValue(dStart))      '11:59 pm - future term
		sTermRelative = "Future"
     ElseIf dEnd >= dtNow Then
        dDisplayDate = (DateValue(dtNow) + CDate("12:00:00 PM"))  '12:00 pm - current term
		sTermRelative = "Current"
     Else: ' term is completely in past
        dDisplayDate = (DateValue(dEnd) + CDate("12:01:00 AM"))   '12:00 am - past term
		sTermRelative = "Past"
     End If
  Else
     dDisplayDate = DateAdd("n", 1439, DateValue(dtNow))
  End If 

	' if we have both a patron key and term key, then load
	' applications and attributes for the patron using the term
	If lPatronKey <> "" AND lTermKey <> "" Then
		Set oHMS = GetClass("HMSDBSrv.ApplicationRead")
		Set rsPatronApplications = oHMS.GetPatronApplication(GetUserSession("StaffToken"), ,lPatronKey,,,,lTermKey)
	
		Set oHMS = GetClass("HMSDBSrv.AttributeRead")
		Set rsPatronAttributes = oHMS.BrowsePatronAttributesAtDate(GetUserSession("StaffToken"), lPatronKey, dDisplayDate)
	
		Set oHMS = GetClass("HMSDBSrv.ContractRead")
		Set rsElements = oHMS.GetContractElements(GetUserSession("StaffToken"), lPatronKey, dStart, dDisplayDate)
		
		Set oHMS = GetClass("HMSDBSrv.PatronRead")
		Set rsPatron = oHMS.GetPatron(GetUserSession("StaffToken"), lPatronKey)
		
		Set oHMS = GetClass("HMSDBSrv.TermRead")
		Set rsTerms = oHMS.GetTerms(GetUserSession("StaffToken"), lTermKey)
		
		lNextTermKey = GetNextTerm(lTermKey)
		
		Set oHMS = GetClass("HMSDBSrv.ApplicationRead")
		Set rsPatronApplications2 = oHMS.GetPatronApplication(GetUserSession("StaffToken"), ,lPatronKey,,,,lNextTermKey)
		
		rsPatronApplications2.Filter = "(Term_Key="& lNextTermKey & " AND " & _
			"SubmittedDate<>Null AND CancelledDate=Null AND ApplicationType_Key=1)"  

		If rsPatronApplications2.EOF Then
			'Student doesn't have app for next term, no reason to hold deposit
			bHasFutureApp = False
		Else
			bHasFutureApp = True
		End If
		
	End If	'consider moving this End If below so everything is wrapped
			'in this if block
	
	' TODO consider a test here or somewhere against if GetCancelledDate()
	' returns false, we shouldn't continue processing...
	
	
	' If they have a future app, override the deposit decision except for
	If bHasFutureApp = True Then 
		sDepositDecision = "Hold deposit. "
		sDecision = sDecision & " Has future application."
	End If
	
	Dim fso_read, fso_write, inFile, outFile, ts, line, text
	Set fso_read = Server.CreateObject("scripting.FileSystemObject")
	Set fso_write = Server.CreateObject("scripting.FileSystemObject")
	
	Set inFile = fso_read.GetFile("D:\OdysseyWeb\HMSWebStaff\tests\TestData.txt")
	Set outFile = fso_write.CreateTextFile("D:\OdysseyWeb\HMSWebStaff\tests\TestOutput.txt")
	Set ts = inFile.OpenAsTextStream(1, -2)
	
	
		
	Do While Not ts.AtEndOfStream
		text = ""
		line = ""
		sDecision = ""
		sDepositDecision = ""
		sBCF = ""
		sAssignmentMessage = ""
		line = ts.ReadLine
		Set oHMS = GetClass("HMSDBSrv.PatronRead")
		Set rsPatron = oHMS.GetPatron(GetUserSession("StaffToken"), ,line)
		lPatronKey = rsPatron("Patron_Key").Value	
		
		sBCF = DetermineBreakContractFee(lPatronKey, lTermKey, sTermRelative, _
									dStart, dDisplayDate)

		ProcessAssignment lPatronKey, lTermKey, sTermRelative, dStart, sAssignmentMessage
	
		ProcessDeposit lPatronKey, lTermKey, GetCancelledDate(lPatronKey, lTermKey), sTermRelative, _
						GetAssignedDate(lPatronKey, dStart), dStart, sDecision, sDepositDecision
		
		text = line & " " & sDecision & " " & sDepositDecision & " BCF: " & sBCF & " " & sAssignmentMessage
		outFile.WriteLine(text)
	Loop

		ts.Close

		
		Set fso_read = Nothing
		Set fso_write = Nothing
	Dim sTitle
	sTitle = oPage.Title

%>

<head><title>Odyssey HMS - <% Response.Write Server.HTMLEncode(sTitle) %> </title>
<!--#include file="include/GlobalInclude.inc" -->
<link rel="stylesheet" type="text/css" media="all" href="css/StaffSite.css" />

<script language="JavaScript" src="script/date.js"></script>
<script language="JavaScript" src="script/AnchorPosition.js"></script>
<script language="JavaScript" src="script/PopupWindow.js"></script>
<script language="JavaScript" src="script/CalendarPopup.js"></script>

<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-us" />
<meta name="ROBOTS" content="ALL" />
<meta name="Copyright" content="Copyright (c) 2004 The CBORD Group, Inc." />

</head>
<body>
<script language="JavaScript">

  cal = new CalendarPopup();
  cal.setWindowProperties('toolbars=no,scrollbars=no,resizable=no');
  cal.showYearNavigation();


function selectItem(command, itemtype, key)
{
  if (itemtype == 'PatronSearch')
    window.location.href = 'PatronCustomForm1.asp?PatronKey=' + key
  else if (itemtype == 'TermSearch')
    window.location.href = 'PatronCustomForm1.asp?TermKey=' + key
  else if (itemtype == 'PatronNewID')
   window.location.href = 'PatronGeneral.asp?cmd=new&id=' + key
}

</script>
 <div id="customDiv1"></div>
 <div id="customDiv2"></div>
 <div id="customDiv3"></div>

<%
  RenderSidebar "tree-PatronCustomForm1"

  Response.Write "<div id=""container"">" & vbCrlf
  Response.Write "<div id=""content"">" & vbCrLf

  RenderPatronHeader sTitle, sPatronTitle

 
  Response.Write (oPage.CustomContent)
 
	Response.write "Testing should be finished now..."
	
  Response.Write "</div></div>" & vbCrLf
  Response.Write "<div id=""footer"">" & vbCrLf & Application("Footer") & vbCrLf & "</div>" & vbCrLf
  Response.Write "</body></html>" & vbCrLf

  Set oFunction = Nothing
  Set oPage = Nothing
  Server.Scriptimeout = 90
%>
