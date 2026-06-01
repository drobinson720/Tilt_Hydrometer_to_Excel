# tiltclass.py
#
# class or sotring tilt data and automating tests

import datetime
import time

DEBUG_MODE = False

class TiltClass:

	lastSeen 	= datetime.datetime(2020, 1, 1) # declare lastSeen and set it arbitrarily to 01/01/2020 so that it has not been seen recently.
	seenTimer	= 180				# time in seconds to have been considered to be seen recently, defaults to 3 minutes

	lastUploaded	= datetime.datetime(2020, 1, 1) # declare lastUploaded and set it arbitrarily to 01/01/2020
	needsUpload	= False				# does this data need uploading to the cloud, defaults to false.
	uploadTimer	= 15				# time in minutes between uploads to the cloud

	tiltTemp	= 0				# declare tilt probe temp as 0, note this is in fahrenheit
	tiltGravity	= "0.000"

	def __init__(self,uuid,name,uploadtime=15):		# uploadtime is the time in MINUTES between planned uploads, defaults to 15 mins
		self.tiltUUID = uuid
		self.tiltName = name
		self.uploadTimer = uploadtime

		if DEBUG_MODE:
			print("*TiltClass Constructor*")
			print(uuid,":",name,":",uploadtime)
		pass

	def checkUUID(self,uuid):
		if self.tiltUUID==uuid:
			return True
		return False

	def checkName(self,name):
		if self.tiltName==name:
			return True
		return False

	def tempCelsius(self):				# convert temperature to celsius and return
		return round((self.tiltTemp / 10 - 32) / 1.8, 1)

	def tempFahrenheit(self):			# return temperature in fahrenheit
		return round(self.tiltTemp / 10, 1)

	def specificGravity(self):			# return specific gravity reading
		return self.tiltGravity / 10000

	def tiltUpdate(self,temp,gravity):		# update values from passed variables, and then set lastSeen to now and needsUpload to true.
                self.tiltTemp		= temp
                self.tiltGravity	= gravity
                self.lastSeen	 	= datetime.datetime.now()

                if (self.timeSinceUpload()>(self.uploadTimer*60)):
                        self.needsUpload=True
                else:
                        self.needsUpload=False

                if DEBUG_MODE:
                        print("name: ",self.tiltName,", temp: ",self.tiltTemp,", sg: ",self.tiltGravity,", upload:", self.needsUpload, " last upload:", self.lastUploaded, " time since upload:",self.timeSinceUpload())
                        
	def timeSinceSeen(self):			# returns number of seconds since last update
		timeDelta = datetime.datetime.now() - self.lastSeen
		timeSeconds = timeDelta.total_seconds()
		return timeSeconds

	def timeSinceUpload(self):			# returns number of seconds since last upload
		timeDelta = datetime.datetime.now() - self.lastUploaded
		timeSeconds = timeDelta.total_seconds()
		return timeSeconds

	def setUpload(self):
		self.lastUploaded = datetime.datetime.now()
		self.needsUpload = False

		if DEBUG_MODE:
			print("classTilt - setUpload: self.lastUpload:", self.lastUploaded)

	def seenRecently(self):
		if (self.timeSinceSeen()<self.seenTimer):
			return True
		return False

	def __str__(self):
		return self.tiltGravity
