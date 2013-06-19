import datetime
import time
import os


class GMT6(datetime.tzinfo):
	def utcoffset(self, dt):
		return datetime.timedelta(hours=-6)

	def tzname(self, dt):
		return "GMT -6"

	def dst(self, dt):
		return datetime.timedelta(0)
