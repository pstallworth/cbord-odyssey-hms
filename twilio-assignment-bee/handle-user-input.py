import win32com.client
import cgi
import os
import time
import datetime
from gmt6 import GMT6
import pythoncom
from pywintypes import Time as PyTime
from twilio import twiml
print ("Content-type: text/html")
print ()

#set up the pytime object for future calls
mytz = GMT6() #have to have a timezone to make aware object

dDateTime = datetime.datetime(
	year = 2013,
	month =8,
	day = 24,
	hour = 10,
	minute = 00,
	second = 00,
	tzinfo=mytz
)

pDateTime = PyTime(dDateTime)

#handle the form processing first
form = cgi.FieldStorage()
PatronID = form.getfirst("Digits", "")

if PatronID == "":
	#handle error here
	pass

#now setup the odyssey web stuff	
hms = win32com.client.Dispatch("HMSDBSrv.SystemRead")
web_user = win32com.client.Dispatch("OdysseyWeb.WebUser")

#may want to create a user for accessing the web api
token = web_user.LoginStaff("username","password", 1)

#this will get us to patron key
sWhere = "(1.-49," + PatronID + ",=)"

#1.-50 is alternate ID (username)
hms = win32com.client.Dispatch("HMSDBSrv.PatronRead")
res = hms.BrowsePatronsGeneral2(token, "1.-50", sWhere)

PatronKey = res.Fields("Patron_Key")

oContract = win32com.client.Dispatch("HMSDBSrv.ContractRead")


rsElements = oContract.GetContractElements(token, PatronKey, pDateTime)
facilityKey = None

# this should get the contract element
# state 5 is expired, state 7 is cancelled
while not rsElements.EOF:
	if rsElements.Fields("Facility_Key") is not None:
		if not (rsElements.Fields("State_ID").Value == 5 or rsElements.Fields("State_ID").Value == 7):
			facilityKey = rsElements.Fields("Facility_Key").Value
	rsElements.MoveNext()

r = twiml.Response()

if facilityKey is None:
	r.say("You do not have an assignment for the fall.")
	r.pause(lenght=1)
	r.say("Goodbye")
	print(str(r))
else:
	facilityRead = win32com.client.Dispatch("HMSDBSrv.FacilityRead")

	rsFacility = facilityRead.GetFacility(
		token, 
		facilityKey, 
		"2.-200", 
		dDateTime, 
		pythoncom.Empty, 
		pythoncom.Empty, 
		pythoncom.Empty
	)

	facilityName = str(rsFacility[0].Fields("2.-200"))
	roomName = str(rsFacility[0].Fields("Name"))

	#most of our facilities have a 'nn - ' at the beginning, so strip that
	if facilityName != 'University Woods' and facilityName != 'Lumberjack Landing':
		facilityName = facilityName[4:]

	r.say("You are assigned to: " + str(facilityName) + " room " + str(roomName))
	r.pause(lenght=2)
	r.say("Goodbye")
	print(str(r))
