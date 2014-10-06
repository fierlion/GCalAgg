################################################
### Google Calendar Aggregator App 
### With Oregon state University Calendar import
### Author: Ray Allan Foote
### Date: Aug 16, 2014
### Version: 0.2
################################################

#imports for calendar interface
import getopt
import sys
import httplib2
from oauth2client.file import Storage
from oauth2client.client import AccessTokenRefreshError
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run
from apiclient.discovery import build
#imports for datetime
from datetime import datetime, timedelta, date
import pytz
import dateutil.parser
#import for bitarray
from bitarray import bitarray
#import for OSU calendar
from bs4 import BeautifulSoup
import mechanize
import os
import re

#google API scpecifics
client_id = ''
client_secret = ''
scope = 'https://www.googleapis.com/auth/calendar'
flow = OAuth2WebServerFlow(client_id, client_secret, scope)
storage = Storage('credentials.dat')
credentials = storage.get()
if credentials is None or credentials.invalid:
  credentials = run(flow, storage)
http = httplib2.Http()
http = credentials.authorize(http)
service = build('calendar', 'v3', http=http)

#OSU calendar specifics
CACHE_DIR = '.cache'
#builds following get request:
#http://catalog.oregonstate.edu/CourseDetail.aspx?subjectcode=CS&coursenumber=325
BASE_URL = 'http://catalog.oregonstate.edu'
SCHEDULE_COURSE_DET = BASE_URL + '/CourseDetail.aspx?'
SUBJECT = 'subjectcode=%s&'
NUMBER = 'coursenumber=%s'

#taking into account daylight savings
OSU_TZ = 'US/Pacific'

#________

#Inputs: days, hours, minutes
#returns time delta object
def mkTD(inD, inH, inM):
  return timedelta(days=inD, hours=inH, minutes=inM)

#inputs string in the form: "2014-08-17T16:00:00", user timezone
#returns datetime object with proper timezone
def datetimify(date_string, usertz):
    fmt = "%Y-%m-%dT%H:%M:%S"
    cest = pytz.timezone(usertz)
    date = datetime.strptime(date_string, fmt).replace(tzinfo=cest)
    return date

#inputs: used together with datetimify
##inputs string in the form: "2014-08-17T16:00:00", user timezone
#returns timedelta object for use in main functions.
def mkStrTD(stringIn, usertz):
  dt = datetimify(stringIn, usertz)
  cest = pytz.timezone(usertz)
  now = datetime.now(tz=cest)
  timeMin = datetime(year=now.year, month=now.month, day=now.day, tzinfo=cest)
  total = dt - timeMin
  return mkTD(total.days, total.seconds//3600, (total.seconds//60)%60)

#Inputs: user email, user timezone (string), start day (int -- 0 for 'now'), end day (int)
#returns two arrays of google formatted (ISO 8601) datetime objects 
#these events represent the start times for all events in start-end window
#and end times for all events in start-end window
def getEvents(userEmail, usertz, startTD, endTD):
  starts = []
  ends = []
  cest = pytz.timezone(usertz)
  now = datetime.now(tz=cest)
  timeMin = datetime(year=now.year, month=now.month, day=now.day, tzinfo=cest) + startTD
  timeMin = timeMin.isoformat()
  timeMax = datetime(year=now.year, month=now.month, day=now.day, tzinfo=cest) + endTD
  timeMax = timeMax.isoformat()
  try:
    request = service.events().list(calendarId=userEmail, timeMin=timeMin, timeMax=timeMax)
    while request:
      response = request.execute()
      for event in response.get('items', []):
        #drop all-day events/un-timed events
        if (event.get('end')):
          evStart = event.get('start')
          evEnd = event.get('end')
          #get start/end time for each event
          dtStart = evStart['dateTime']
          dtEnd = evEnd['dateTime']
          starts.append(dtStart)
          ends.append(dtEnd)
      request = service.events().list_next(request, response)
  except AccessTokenRefreshError:
    print ('The credentials have been revoked or expired, please re-run'
           'the application to re-authorize')
  return starts, ends
 
#Inputs: two arrays of datetime objects: one for start times, one for end times
#also an int for start time in (should be the same start time as getEvents, taken
# from the user input time window)
#usertimezone is still tricky for events around midnight/1am due to complications 
#with Daylight savings.
def getBusy(startArr, endArr, startTD, usertz):
  starts = []
  ends = []
  #set initial comparator to midnight of day one
  cest = pytz.timezone(usertz)
  now = datetime.now(tz=cest)
  timeMin = datetime(year=now.year, month=now.month, day=now.day, tzinfo=cest) + startTD
  #helper to determine > 24 hours or not
  one_day = timedelta(days=1)
  startDay = startTD.days
  #parse dt objects
  for date in startArr:
    inner = []
    dtstart = dateutil.parser.parse(date)
    if((dtstart - timeMin) >= one_day):
      days = dtstart - timeMin
      inner.append(days.days + startDay)
      inner.append(dtstart.hour)
      inner.append(dtstart.minute)
    else:
      inner.append(0 + startDay)
      inner.append(dtstart.hour)
      inner.append(dtstart.minute)
    starts.append(inner)
  for date in endArr:
    inner = []
    dtend = dateutil.parser.parse(date)
    if((dtend - timeMin) >= one_day):
      days = dtend - timeMin
      inner.append(days.days + startDay)
      inner.append(dtend.hour)
      inner.append(dtend.minute)
    else:
      inner.append(0 + startDay)
      inner.append(dtend.hour)
      inner.append(dtend.minute)
    ends.append(inner)
  return starts, ends


#inputs: two arrays startArr and endArr, containing [d, h, m] int arrays.  
#outputs: a bitarray object initially all 0, then with 1s set from every start to every end, 
#bit array object is the size from 0-endTD / in 30 minute increments.
def bitify(startArr, endArr, endTD):
  startBit = []
  endBit = []
  #number of bits in array should be number days from 12am present day - endIn * 48
  #this makes bit array for half-hour increments.
  numBits = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  #next two lines are hack to make it work
  if numBits < 50000:
    numBits = 50000
  myArray = bitarray(numBits)
  myArray.setall(0)
  #get first event for comparison and make datetime object
  for i in range (0, len(startArr)):
    #get bits for start/end time
    if startArr[i][0] == 0:
      startBit.append((startArr[i][1] * 2) + startArr[i][2]/30)
    else:
      startBit.append(((startArr[i][1] * 2) + startArr[i][2]/30) + startArr[i][0] * 48)
  for i in range (0, len(endArr)):
    if endArr[i][0] == 0:
      endBit.append((endArr[i][1] * 2) + endArr[i][2]/30)
    else:
      endBit.append(((endArr[i][1] * 2) + endArr[i][2]/30) + endArr[i][0] * 48)
  #set bits in bitArray
  for i in range(0, len(startBit)):
    for j in range (startBit[i], endBit[i]):
      myArray[j] = 1
  return myArray


#helper function
def timeWinArr(subStartTD, subEndTD, endTD):
  numBits = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  #next two lines are hack to make it work
  if numBits < 50000:
    numBits = 50000
  myArray = bitarray(numBits)
  myArray.setall(1)
  startBit = 0
  endBit = 0
  #get bits for start/end time
  if subStartTD.days == 0:
    startBit = (subStartTD.seconds//3600 * 2) + (((subStartTD.seconds//60)%60)/30)
  else:
    startBit = (subStartTD.seconds//3600 * 2) + (((subStartTD.seconds//60)%60)/30) + (subStartTD.days * 48)
  if subEndTD.days == 0:
    endBit = (subEndTD.seconds//3600 * 2) + (((subEndTD.seconds//60)%60)/30)
  else:
    endBit = (subEndTD.seconds//3600 * 2) + (((subEndTD.seconds//60)%60)/30) + (subEndTD.days * 48)
    #set bits in bitArray
  for i in range(startBit, endBit):
    myArray[i] = 0
  return myArray

#main function for free times.  Use with stringTimes to get formatted result array.
#inputs: email array with [(contactEmail, contactTimezone),(etc.)] 
#timeDelta objects for start/end times.  
def getFreeTimes(emailArrIn, startSTR, endSTR, USER_TZ):
  startTD = mkStrTD(startSTR, USER_TZ)
  endTD = mkStrTD(endSTR, USER_TZ)
  #initialize result array to total num of bits (endIn * 48)
  startBit = (startTD.days * 48) + (startTD.seconds//3600 * 2) + (((startTD.seconds//60)%60)/30)
  endBit = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  numBits = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  #next two lines are hack to make it work
  if numBits < 50000:
    numBits = 50000
  resultArr = bitarray(numBits)
  resultArr.setall(0)
  #get events in default 48 hour window (0-2)
  for email in emailArrIn:
    userEvents = getEvents(email[0], email[1], startTD, endTD)
    userBits = getBusy(userEvents[0], userEvents[1], startTD, email[1])
    userArray = bitify(userBits[0], userBits[1], endTD)
    resultArr |= userArray
  return resultArr

#inputs: bit array with true for each time to be printed
#start/end timedelta
#default user timezone.  
#output: formatted strings of all true times in bit array
def stringTimes(bitArIn, startSTR, endSTR, USER_TZ):
  startTD = mkStrTD(startSTR, USER_TZ)
  endTD = mkStrTD(endSTR, USER_TZ)
  timesArr = []
  bits = bitArIn
  startBit = (startTD.days * 48) + (startTD.seconds//3600 * 2) + (((startTD.seconds//60)%60)/30)
  endBit = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  for bit in range (startBit, endBit):
    if bits[bit] == 0:
      stringed = ""
      day = bit / 48
      hour = (bit % 48) / 2
      if bit % 2 == 0:
        minute = 0
      else:
        minute = 30
      #"Aug 8, 2014, 4:30PM - Aug 14, 2014, 4:30PM"
      freeTime = mkTD(day, hour, minute)
      cest = pytz.timezone(USER_TZ)
      now = datetime.now(tz=cest)
      timeMin = datetime(year=now.year, month=now.month, day=now.day, tzinfo=cest)
      newDate = timeMin+freeTime
      t = newDate.time()
      stringed = "{0}, {1}-{2}-{3}".format(str(t)[0:5], str(newDate.day), str(newDate.month), str(newDate.year))
      timesArr.append(stringed)
  return timesArr

#main function for free names.  
#inputs: email array with [(contactEmail, contactTimezone),(etc.)] 
#iso formatted strings, tzinfo for window.
def getFreeNames(emailArrIn, startSTR, endSTR, USER_TZ, OSU_ARR):
  startTD = mkStrTD(startSTR, USER_TZ)
  endTD = mkStrTD(endSTR, USER_TZ)
  names = []
  newArr = timeWinArr(startTD, endTD, endTD)
  numBits = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  #next two lines are hack to make it work
  if numBits < 50000:
    numBits = 50000
  resultArr = bitarray(numBits)  
  resultArr.setall(0)
  resultArr |= newArr
  for email in emailArrIn:
    userEvents = getEvents(email[0], email[1], startTD, endTD)
    userBits = getBusy(userEvents[0], userEvents[1], startTD, email[1])
    userArray = bitify(userBits[0], userBits[1], endTD)
    resultArr |= userArray
    if OSU_ARR != 0:
      resultArr |= OSU_ARR
    if (resultArr == newArr):
      names.append(email[0])
    resultArr.setall(0)
    resultArr |= newArr
  return names

#input: url for page to scrape
#output: string of html code
def getPage(url):
  br = mechanize.Browser()
  br.set_handle_refresh(False)
  br.set_handle_robots(False)
  br.addheaders
  response = br.open(url)
  return str(response.read())

#input: dept/number/courseProf in form "CS", "419", "McGrath, D." 
#Each MUST be in this format (courseProf is a regex)
#if TBA, returns empty.
def getDetails(sub, num, prof):
  courseSub = sub
  courseNum = num
  courseProf = prof
  courseArray = []
  details = []
  url = SCHEDULE_COURSE_DET + (SUBJECT % courseSub) + (NUMBER % courseNum)
  htmlDoc = getPage(url)
  soup = BeautifulSoup(htmlDoc)
  classes = soup.find_all("td", text=courseProf)
  for link in classes:
    courseArray.append(link.find_next_sibling("td").get_text().strip())
  for course in courseArray:
    if course != "TBA":
      space = course.index(' ')
      hyphen = course.index('-', space+10)
      mid = []
      string = str(course[:space])
      mid.append(list(string))
      mid.append(str(course[space+1:space+5]))
      mid.append(str(course[space+6:space+10]))
      mid.append(str(course[space+10:hyphen]))
      mid.append(str(course[hyphen+1:]))
      details.append(mid)
  return details

#input: string array from getDetails.
#strings in format: [ [['R'], '1600', '1750', '3/30/15', '6/5/15'] , [['M','T']...etc] ]
#output: datetime objects for start and end time.
#function expects that startDate is correct (['R'] == '3/30/15')
def datify(deetsIn):
  evArr = []
  if len(deetsIn) > 0:
    for deet in deetsIn:
      inArr = []
      #dayArr = deet[0]
      startMin = int(deet[1][(len(deet[1]))-2:])
      startHour = int(deet[1][:(len(deet[1]))-2])
      #round start down
      if startMin > 0:
        if startMin >= 30:
          startMin = 30
        startMin = 0
      endMin = int(deet[2][(len(deet[2]))-2:])
      endHour = int(deet[2][:(len(deet[2]))-2])
      if endMin > 0:
        if endMin > 30:
          endMin = 0
          endHour += 1
        else:
          endMin = 30
      splitStart = deet[3].split('/')
      startYear = 2000+int(splitStart[2])
      startMonth = int(splitStart[0])
      startDay = int(splitStart[1])
      startDaS = datetime(startYear, startMonth, startDay, startHour, startMin)
      startDaE = datetime(startYear, startMonth, startDay, endHour, endMin)
      splitEnd = deet[4].split('/')
      endYear = 2000+int(splitEnd[2])
      endMonth = int(splitEnd[0])
      endDay = int(splitEnd[1])
      endDate = datetime(endYear, endMonth, endDay, endHour, endMin)
      inArr.append(deet[0])
      inArr.append(startDaS)
      inArr.append(startDaE)
      inArr.append(endDate)
      evArr.append(inArr)
  return evArr

#input: string array from getDetails.
#[['M', 'W', 'F'], datetime.datetime(2015, 1, 5, 12, 0), datetime.datetime(2015, 3, 13, 12, 50)]
def osuBusy(evArrIn, OSU_TZ, startIn):
  dayDict = {'M':1,'T':2,'W':3,'R':4,'F':5}
  cest = pytz.timezone(OSU_TZ)
  now = datetime.now(tz=cest)
  timeMin = datetime(year=now.year, month=now.month, day=now.day, tzinfo=cest) + startIn
  startArr = []
  endArr = []
  cest = pytz.timezone(OSU_TZ)
  aWeek = timedelta(days=7)
  if len(evArrIn) > 0:
    for event in evArrIn:
      dateS = event[1].replace(tzinfo=cest)
      dateE = event[2].replace(tzinfo=cest)
      finalDate = event[3].replace(tzinfo=cest)
      dayArray = event[0]
      dayArrNums = []
      for day in dayArray:
        dayArrNums.append(dayDict[day])
      dayOne = dayArrNums.pop(0)
      #check that date is after min date
      if dateS >= timeMin:
        while dateS < finalDate:    
          startArr.append(dateS.isoformat())
          endArr.append(dateE.isoformat())
          for num in dayArrNums:
            startIn = (dateS + timedelta(days=(num - dayOne)))
            startArr.append(startIn.isoformat())
            endIn = (dateE + timedelta(days=(num - dayOne)))
            endArr.append(endIn.isoformat())
          dateS += aWeek
          dateE += aWeek
  return startArr, endArr

#main function for osu times.  Use with stringTimes to get formatted result array.
#inputs: array with [(dept/number/courseProf),(d,n,c)...] each in form "CS", "419", "McGrath, D."
#iso formatted strings, tzinfo for window.
#output: formatted array
def getOSUTimes(classArrIn, startSTR, endSTR, USER_TZ):
  startTD = mkStrTD(startSTR, USER_TZ)
  endTD = mkStrTD(endSTR, USER_TZ)
  startBit = (startTD.days * 48) + (startTD.seconds//3600 * 2) + (((startTD.seconds//60)%60)/30)
  endBit = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  numBits = (endTD.days * 48) + (endTD.seconds//3600 * 2) + (((endTD.seconds//60)%60)/30)
  #next two lines are hack to make it work
  if numBits < 50000:
    numBits = 50000
  resultArr = bitarray(numBits)
  resultArr.setall(0)
  #get events in default 48 hour window (0-2)
  for classIn in classArrIn:
    details = getDetails(classIn[0], classIn[1], classIn[2])
    dates = datify(details)
    full = osuBusy(dates, OSU_TZ, startTD)
    fullBits = getBusy(full[0], full[1], startTD, OSU_TZ)
    fullArray = bitify(fullBits[0], fullBits[1], endTD)
    resultArr |= fullArray
  return resultArr

#use to divide up imput strings
#outer array divides with '~', inner array divides with '|'
def parseInputString(inputStr):
  outArr = inputStr.split('~')
  splitArr = []
  for string in outArr:
    if len(string.split('|')) != 1:
      splitArr.append(string.split('|'))
    else:
      return outArr
  return splitArr

def usage():
    print "To use: ./calOSUcom.py -tewo [args in separated by '~', inner args separated by '|']\n"

#_______________

def main():
  USER_TZIn = ''
  emailTZArrIn = ''
  timeIn = ''
  osuArraIn = ''

  #getopts.  To change useage info, modify useage() above
  #use:
#  ./calOSUcom.py -t US/Mountain 
#-e 'fierlion@gmail.com|US/Mountain~footer@onid.oregonstate.edu|US/Mountain~lawinge@onid.oregonstate.edu|US/Central' 
#-w 2014-09-29T09:00:00~2014-09-29T20:00:00 
#-o 'CS|419|Zhang, E.~CS|419|Ramsey, S.~CS|419|Bobba, R.'
#-a/-b
  try:
    opts, args = getopt.getopt(sys.argv[1:], "ht:e:w:o:ab", ["help"])
  except getopt.GetoptError as err:
    print "For use information type -h or --help\n"
    print(err)
    usage()
    sys.exit(2)

  for o, a in opts:
    #collect data
    if o in ("-t"):
      USER_TZIn = a
    elif o in ("-e"):
      emailTZArrIn = a
    elif o in ("-w"):
      timeIn = a
    elif o in ("-o"):
      osuArraIn = a
    
      #use cases 
    elif o in ("-a"):
      #use case 1: 
      #inputs: usertz, email array, window, osu data(optional)
      #output:all free times
      USER_TZ = USER_TZIn
      emailTZArr = parseInputString(emailTZArrIn)
      START_ARR = parseInputString(timeIn) 
      START_STR = str(START_ARR[0])
      END_STR = str(START_ARR[1])
      timeArray = getFreeTimes(emailTZArr, START_STR, END_STR, USER_TZ)
      if osuArraIn != '':
        osuArr = parseInputString(osuArraIn)
        classArray = getOSUTimes(osuArr, START_STR, END_STR, USER_TZ)
        timeArray |= classArray
      resultTimes = stringTimes(timeArray, START_STR, END_STR, USER_TZ)
      for time in resultTimes:
        print time

    elif o in ("-b"):
      #use case 2: 
      #inputs: usertz, email array, window, osu data(optional)
      #output:all free names
      USER_TZ = USER_TZIn
      emailTZArr = parseInputString(emailTZArrIn)
      START_ARR = parseInputString(timeIn) 
      START_STR = str(START_ARR[0])
      END_STR = str(START_ARR[1])
      if osuArraIn != '':
        osuArr = parseInputString(osuArraIn)
        classArray = getOSUTimes(osuArr, START_STR, END_STR, USER_TZ)
        resultNames = getFreeNames(emailTZArr, START_STR, END_STR, USER_TZ, classArray)
      resultNames = getFreeNames(emailTZArr, START_STR, END_STR, USER_TZ, OSU_ARR = 0)
      for name in resultNames:
        print name
    elif o in ("-h", "--help"):
      usage()
      sys.exit()
    else:
      assert False, "unhandled option"


if __name__ == '__main__':
  main()
