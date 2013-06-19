import cgi
import cgitb
import datetime
import os
from pywintypes import Time as PyTime
import time
from twilio import twiml
import win32com.client

cgitb.enable()

print ("Content-type: text/xml")
print ()

r = twiml.Response()

#get patron ID from phone input
r.say("Please enter your student ID number")
r.gather(numDigits="8", action="handle-user-input.py", method="POST")

print (str(r))