# Script to convert json to xml and dump it to a directory.
import self as self
from icalendar import vCalAddress
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from json2xml import json2xml
import json as j
import shutil
import os

#import email.MIMEText
#import email.MIMEBase
#from email.MIMEMultipart import MIMEMultipart
import email.mime.text
import email.mime.base
import email.mime.multipart
import smtplib
import datetime as dt
import icalendar
import pytz
from pathlib import Path

def inbox_path():
    return ".\inbox"

def sap_bucket():
    return ".\sap_bucket"

def processed_path():
    return ".\processed"

def process_for_outlook(filepath, filename):
    with open(filepath, "r") as file:
        filedata = file.read()
        jsonstring = j.loads(filedata)

    print(jsonstring)

    eventList = []
    for events in jsonstring:
        event = {}
        for key, val in events.items():
            event[key] = val
            #print(key, val)
            
        eventList.append(event)

def process_for_sap(filepath, filename):
    # Use a breakpoint in the code line below to debug your script.
    # Press Ctrl+F8 to toggle the breakpoint.
    with open(filepath, "r") as file:
        filedata = file.read()
        jsonstring = j.loads(filedata)

    print(jsonstring)

    # Convert json to xml
    xmlstring = json2xml.Json2xml(jsonstring).to_xml()
    print(xmlstring)

    # Write to xml file
    xmlfilename = filename.replace('.txt', '.xml')
    xmlfile = open(xmlfilename, 'w')
    xmlfile.write(xmlstring)
    xmlfile.close()

    # Move xml file to SAP bucket
    outputpath = os.path.join(sap_bucket(), xmlfilename)
    shutil.move(xmlfilename, outputpath)

    # Move json file to processed folder
    processedpath = os.path.join(processed_path(), filename)
    shutil.move(filepath, processedpath)

# Imagine this function is part of a class which provides the necessary config data
# https://learnpython.com/blog/working-with-icalendar-with-python/
# https://learnpython.com/blog/working-with-icalendar-with-python/
#def send_appointment(self, attendee_email, organiser_email, subj, description, location, start_hour, start_minute):
def send_invite():
  # Timezone to use for our dates - change as needed
  tz = pytz.timezone("Europe/London")
  #start = tz.localize(dt.datetime.combine(self.date, dt.time(start_hour, start_minute, 0)))
  # Build the event itself
  cal = icalendar.Calendar()
  cal.add('prodid', '-//Unclass calendar event //example.com//')
  cal.add('version', '2.0')
  cal.add('method', "REQUEST")
  event = icalendar.Event()
  attendee_email = vCalAddress('MAILTO:rdoe@example.com')
  event.add('attendee', attendee_email, encode=0)
  organiser_email = vCalAddress('MAILTO:service@example.com')
  event.add('organizer', organiser_email)
  # event.add('attendee', attendee_email)
  # event.add('organizer', organiser_email)
  event.add('status', "confirmed")
  event.add('category', "Event")
  event.add('summary', "Meeting with vendor")
  #event.add('summary', subj)
  event.add('description', "Meeting with vendor")
  #event.add('description', description)
  event.add('location',"Jurong East")
  #event.add('location', location)
  event.add('dtstart', dt.datetime(2022, 1, 25, 8, 0, 0, tzinfo=pytz.utc))
  event.add('dtend', dt.datetime(2022, 1, 25, 10, 0, 0, tzinfo=pytz.utc))
  #event.add('dtstart', start)
  #event.add('dtend', tz.localize(dt.datetime.combine(self.date, dt.time(start_hour + 1, start_minute, 0))))
  #event.add('dtstamp', tz.localize(dt.datetime.combine(self.date, dt.time(6, 0, 0))))
  #event['uid'] = self.get_unique_id() # Generate some unique ID
  event['uid'] = 1234
  event.add('priority', 5)
  event.add('sequence', 1)
  event.add('created', tz.localize(dt.datetime.now()))

  # Add a reminder
  #alarm = icalendar.Alarm()
  #alarm.add("action", "DISPLAY")
  #alarm.add('description', "Reminder")
  # The only way to convince Outlook to do it correctly
  #alarm.add("TRIGGER;RELATED=START", "-PT{0}H".format(reminder_hours))
  #event.add_component(alarm)

  cal.add_component(event)

  # Build the email message and attach the event to it
  msg = email.mime.multipart.MIMEMultipart("alternative")
  #msg["Subject"] = subj
  msg["Subject"] = "Meeting with vendor"
  msg["From"] = organiser_email
  msg["To"] = attendee_email
  msg["Content-class"] = "urn:content-classes:calendarmessage"

  #msg.attach(email.MIMEText.MIMEText(description))
  msg.attach(email.mime.text.MIMEText("Meeting with vendor"))

  filename = "invite.ics"
  part = email.mime.base.MIMEBase('text', "calendar", method="REQUEST", name=filename)
  part.set_payload( cal.to_ical() )
  #email.encoders.encode_base64(part)
  part.add_header('Content-Description', filename)
  part.add_header("Content-class", "urn:content-classes:calendarmessage")
  part.add_header("Filename", filename)
  part.add_header("Path", filename)
  msg.attach(part)

  # Send the email out
  #s = smtplib.SMTP('localhost')
  #s.sendmail(msg["From"], [msg["To"]], msg.as_string())
  #s.quit()

  # Write to disk
  directory = Path.cwd() / 'MyCalendar'
  try:
      directory.mkdir(parents=True, exist_ok=False)
  except FileExistsError:
      print("Folder already exists")
  else:
      print("Folder was created")

  f = open(os.path.join(directory, 'example.ics'), 'wb')
  f.write(cal.to_ical())
  f.close()

  e = open('MyCalendar/example.ics', 'rb')
  ecal = icalendar.Calendar.from_ical(e.read())
  for component in ecal.walk():
      print(component.name)
      if component.name == "VEVENT":
          #print(component.get("name"))
          print(component.get("description"))
          print(component.get("summary"))
          print(component.get("location"))
          print(component.decoded("dtstart"))
          print(component.decoded("dtend"))
  e.close()

# Main function
if __name__ == '__main__':
    # iterate over incoming json files directory
    for filename in os.listdir(inbox_path()):
        filePath = os.path.join(inbox_path(), filename)
        # checking if it is a file
        if os.path.isfile(filePath):
            print(filePath)
            #process_for_sap(filePath,filename)
            process_for_outlook(filePath,filename)

    send_invite()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
