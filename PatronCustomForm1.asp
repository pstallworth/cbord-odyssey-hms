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
' Author: Paul Stallworth pestallworth@sfasu.edu
' Custom Form for processing Cancellations
' This was created to take over the processing of cancellations, and
' replace the use of paper checklists with this form.  I created the
' three main functions for handling the deposit, assignment, and BCf
' based upon the 2012-2013 Residence Hall Contract, specifically sections
' 17 and 18 which have the most detail about cancellations and refunds.

' As a foundation, elements of PatronGeneral.asp were used in addition to
' parts of the default PatronCustomForm1.asp.

  ValidateUser "PatronCustomForm1.asp"

  Dim oPage, oFunction, patronAttributeGroups, groupID
  Set oPage = New CPage
  oPage.InitializeByName "PatronCustomForm1"
  Set oFunction = oPage.PageFunction

  Call PageCodeHeaderPatronCustomForm1(oPage.Key)

  ' verify user has access to patrons
  If GetUserSession("UR_Read_PS") <> 1 Then
    Response.Redirect "Error.asp?Error=" & Server.URLEncode("Access Denied to " & GetPatronLabel() & ". Please see your administrator.")
  End If
  
	' Verify user has access to page.  Access is controlled through
	' the Processing attribute group.  If a user, or user group has the 
	' Processing group in the No Access section of the Exclusive filters for
	' Patron Attributes, they will not be able to access this page.
	patronAttributeGroups = Array("")
	patronAttributeGroups = split(GetUserSession("UR_Filter_No_hfPATS_AttributeGroup"),",")
  
	For Each groupID In patronAttributeGroups
		If groupID = "20" Then
			 Response.Redirect "Error.asp?Error=" & Server.URLEncode("Access Denied to Cancellation Form. Please see your administrator.")
		End if
	Next
  ' If PatronKey QueryString specified, load the specified patron if not already loaded
	Dim lPatronKey, sAssignmentMessage, sCmd, sErrors, sPatronName, sPatronID, bNoteAlert, sPatronTitle, sAltID
	Dim dStart, dEnd, lTermKey, dDisplayDate, dtNow
	Dim oHMS,rsPatronApplications2,rsElements
	Dim dCancelledDate, dSubmittedDate 
	Dim sDepositDecision, sBCF, bCancelled
	Dim lNextTermKey
	Dim sDecision, sTermRelative, bHasFutureApp, dAssignedDate
	Dim dActualStart, bProcessed
	Dim oContractWrite, dContractEnd

	lTermKey = 0
	sErrors = Request.QueryString("Error")
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

	If lTermKey = 0 Then
		Response.Redirect "PatronGeneral.asp?Error=" & Server.URLEncode("Please select a term first.")
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

  
	If Request.QueryString("date1") <> "" Then
		dContractEnd = (DateValue(Request.QueryString("date1")) + Time)
	End If
	
	Dim bInventorySheet
	
	If Request.Form("inventory") = "1" Then
		bInventorySheet = True
	Else
		bInventorySheet = False
	End If
	
	' if we have both a patron key and term key, then load
	' applications and attributes for the patron using the term
	If lPatronKey <> "" AND lTermKey <> "" Then	
		Set oHMS = GetClass("HMSDBSrv.ContractRead")
		Set rsElements = oHMS.GetContractElements(GetUserSession("StaffToken"), lPatronKey, dStart, dDisplayDate)
	
		bHasFutureApp = hasFutureApp(lPatronKey, lTermKey)
		
	End If	'consider moving this End If below so everything is wrapped
			'in this if block

	Dim text
	text = Array("")

	bProcessed = getProcessedStatus(lPatronKey, dStart)
	If bProcessed = False Then
		If Not IsNull(GetCancelledDate(lPatronKey, lTermKey))  Then	
			bCancelled = True
			sBCF = DetermineBreakContractFee(lPatronKey, lTermKey, sTermRelative, _
											dStart, dDisplayDate)

			'ProcessAssignment lPatronKey, lTermKey, sTermRelative, dStart, dDisplayDate, sAssignmentMessage
			ProcessAssignment lPatronKey, lTermKey, sTermRelative, dStart, dEnd, sAssignmentMessage
			
			text = Split(sAssignmentMessage)
				
			DetermineDeposit lPatronKey, lTermKey, GetCancelledDate(lPatronKey, lTermKey), sTermRelative, _
							GetSubmittedDate(lPatronKey, lTermKey), dStart, sDecision, sDepositDecision
							
			' If they have a future app, override the deposit decision
			If bHasFutureApp = True Then 
				sDepositDecision = "Hold."
				sDecision = sDecision & " Has future application."
			End If
		Else
			bCancelled = False
			sDecision = "No cancellation for this semester."
		End If
		
		If dContractEnd <> "" Then
			Dim oContractRead, rsCElements, lElementKey, vFailed, bRet
			Set oContractRead = GetClass("HMSDBSrv.ContractRead")
			Set rsCElements = oContractRead.GetContractElements(GetUserSession("StaffToken"), lPatronKey, dStart, dEnd)

			Dim fName
			If rsCElements.EOF = False Then 
				Do Until rsCElements.EOF
					If Not IsNull(rsCElements("Facility_Key")) Then
						If (rsCElements("State_ID") = 2) Then
							lElementKey = rsCElements("Element_Key")
							fName = rsCElements("Name")
							Exit Do
						End If
					End If
					rsCElements.MoveNext
				Loop
			End If
			
			If lElementKey <> 0 Then
				Set oContractWrite = GetClass("HMSDBSrv.ContractWrite")
				bRet = oContractWrite.UpdateOneContractElement2(GetUserSession("StaffToken"), lElementKey, vFailed,,,4,,,dContractEnd)
				
				If bRet = False Then
					Response.Redirect "Error.asp?Error=Update Contract Element Failed:" & fName & ": " & GetFailedError(vFailed)
				Else
					Response.Redirect "PatronCustomForm1.asp"
				End If
			Else
				Response.Redirect "Error.asp?Error=Could not find active assignment to complete"
			End If
		End If
		
		If Request.Form("apply") = "1" Then
		
			bRet = ApplyDecision(lPatronKey, sBCF, sAssignmentMessage, sDecision, _
						sDepositDecision, dStart, dEnd, bInventorySheet, vFailed) 
			
			
			If bRet = False Then
				Response.Redirect "Error.asp?Error='Error applying decision: ' " & GetFailedError(vFailed)
			Else
				Response.Redirect "PatronCustomForm1.asp"
			End If
			
			' Insert attribute for processing
		End If
	Else
		sDecision = "Cancellation has been processed."
	End If 'end if has been processed	
	
	Dim sTitle
	sTitle = oPage.Title

%>

<head><title>Odyssey HMS - <% Response.Write Server.HTMLEncode(sTitle) %> </title>
<!--#include file="include/GlobalInclude.inc" -->
<link rel="stylesheet" type="text/css" media="all" href="css/StaffSite.css" />
<link rel="stylesheet" type="text/css" media="screen" href="css/SFACustom.css" />
<link rel="stylesheet" type="text/css" media="print" href="css/print.css">
<script language="JavaScript" src="script/date.js"></script>
<script language="JavaScript" src="script/AnchorPosition.js"></script>
<script language="JavaScript" src="script/PopupWindow.js"></script>
<script language="JavaScript" src="script/CalendarPopup.js"></script>

<script language="JavaScript" src="script/myScript.js"></script>
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
	Call PageCodeContentPatronCustomForm1(oPage.key)

	Dim ActualStart, ActualEnd, pDate

If bProcessed = True Then
	Response.write "<p id=""decision"">" & sDecision & "</p>"
ElseIf bCancelled = True Then	
	Response.write "<br />" & vbCrLf & vbCrLf
	Response.write "<ul class=""decision"">" & vbCrLf
	Response.write "<li><b>Primary Decision Indicator:</b> " & sDecision & "</li><br />" & vbCrLf
	Response.write "<li><b>Deposit:</b> " & sDepositDecision & "</li><br />" & vbCrLf
	Response.write "<li><b>BCF:</b> " & sBCF & "</li><br />" & vbCrLf
	Response.write "<li><b>Assignment Decision:</b> " & sAssignmentMessage & "</li><br />" & vbCrLf
	Response.write "</ul>" & vbCrLf


	If text(0) = "Complete" AND Request.QueryString("date1") = "" Then
		Response.write "<div class=""showCheckout"">"  & vbCrLf
	Else
		Response.write "<div class=""hideCheckout"">" & vbCrLf
	End If
		
	Response.write "<form name=""date_select"" action=""PatronCustomForm1.asp"">" & vbCrLf
	Response.write "Checkout date on inventory sheet: " & vbCrLf
	Response.write "<input type=""text"" name=""date1"" value="""" size=12 />" & vbCrLf
	Response.write "<a href=""#"" class=""button_style"" onClick=""cal.select(document.forms['date_select'].date1,'anchor1','MM/dd/yyyy'); return false;"""
	Response.write "NAME=""anchor1"" ID=""anchor1"">Select Date</a>" & vbCrLf
	Response.write "<input type=""hidden"" name=""update"" value=""1"">"
	Response.write "<input type=""submit"" value='Update Contract'>" & vbCrLf
	Response.write "</form>" & vbCrLf
	Response.write "</div>" & vbCrLf
	Response.write "<br />" & vbCrLf
	
	If sDepositDecision = "Refund." OR sDepositDecision = "Forfeit." _
		OR sDepositDecision = "Hold." Then
		Response.Write "<div class=""showApplyDecision"">" & vbCrLf
	Else
		Response.write "<div class=""hideApplyDecision"">" & vbCrLf
	End If
	
	Response.write "<form name=""apply_decision"" action=""PatronCustomForm1.asp"" method=""POST"">" & vbCrLf
	Response.write "<input type=""hidden"" name=""apply"" value=""1"">" & vbCrLf
	Response.write "<input type=""checkbox"" class=""inventory"" name=""inventory"" value=""1""> Received inventory sheet<br />" & vbCrLf
	Response.write "<input type=""submit"" class=""apply_decisions"" value=""Apply Decisions"">" & vbCrLf
	Response.write "</form>" & vbCrLf
	Response.write "</div>" & vbCrlf & vbCrLf
Else
	Response.write "<p id=""decision"">" & sDecision & "</p>"
End if
	Response.write "<br /><br /><a href='#' class=""debug"" onClick='toggleShow()'>Decision Information</a>" & vbCrLf & vbCrLf
	'Decision Criteria
	Response.write "<div id=""debug""><br />" & vbCrLf & vbCrLf
	Response.write "Term Indicator: " & sTermRelative & "<br />" & vbCrLf
	'This line may be able to go
	Response.write "Display Date: " & dDisplayDate & "<br />" & vbCrLf
	Response.write "App Submitted Date: " & GetSubmittedDate(lPatronKey, lTermKey) & "<br />"
	Response.write "Cancellation Deadline: " & GetDeadline(lTermKey) & "<br />" & vbCrLf
	Response.write "Cancelled Date: " & GetCancelledDate(lPatronKey, lTermKey) & "<br />" & vbCrLf
	GetContractActualDates lPatronKey, dStart, dEnd, ActualStart, ActualEnd
	Response.write "Actual Start Date: " & ActualStart & "<br /> Actual End Date: " & ActualEnd & "<br />" & vbCrLf
	Response.write "Academic Suspension: " & IsSuspended(lPatronKey, dStart) & "<br />" & vbCrLf
	Response.write "Discipline Suspension: " & IsDisciplineSuspension(lPatronKey, dDisplayDate) & "<br />" & vbCrLf
	Response.write "Admitted: " & (Not IsRejected(lPatronKey, dStart)) & "<br />" & vbCrLf
	Response.write "Off-Campus Permit: " & getOCPermitStatus(lPatronKey, dStart, ,pDate) & "<br />" & vbCrLf
	Response.write "Permit Date: " & pDate & "<br />" & vbCrLf
	Response.write "Is Assigned: " & IsAssigned(lPatronKey, dStart, dEnd, "") & "<br />" & vbCrLf
	Response.write "Is Enrolled: " & IsEnrolled(lPatronKey, dStart) & "<br />" & vbCrLf
	Response.write "</div>" & vbCrLf
	
	Response.Write "</div></div>" & vbCrLf
	Response.Write "<div id=""footer"">" & vbCrLf & Application("Footer") & vbCrLf & "</div>" & vbCrLf
	Response.write "<div id=""custom_footer"">" & vbCrLf & "Processed by: " 
	
	Response.write GetUserSession("StaffName") & vbCrLf & "</div>" & vbCrLf
	

	Response.Write "</body></html>" & vbCrLf
	Set oFunction = Nothing
	Set oPage = Nothing
%>
