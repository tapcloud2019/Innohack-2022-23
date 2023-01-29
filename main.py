# Script to process incoming JSON files from Powerapps
from dateutil import parser
from icalendar import vCalAddress
import xml.etree.ElementTree as ET
from json2xml import json2xml
import json as j
import shutil
import os
import email.mime.text
import email.mime.base
import email.mime.multipart
import datetime as dt
import icalendar
import pytz
from pathlib import Path

def inbox_path():
    return ".\inbox"

def sap_bucket():
    return ".\sap_bucket"

def processed_sap_path():
    return ".\processed-sap"

def processed_cal_path():
    return ".\processed-cal"

def handle_xml_type(filepath, filename):

    if "LEAVE" in filename:
        print("\tLEAVE json detected")
        process_leave(filepath, filename)
        process_for_sap(filepath, filename, "LEAVE")
    elif "DOTS" in filename:
        print("\tDOTS json detected")
        process_dots(filepath, filename)
        process_for_sap(filepath, filename, "DOTS")
    elif "CALENDAR" in filename:
        print("\tCALENDAR json detected")
        process_for_calendar(filePath, filename)

def process_leave(filepath, filename):
    with open(filepath, "r") as file:
        jsonstring = j.load(file)

    # add pf from filename format: pf,email,datetime,(CloudHR)_TYPE.TXT
    jsonstring[0]["pf"] = filename.split(",")[0]

    newxmltag = "pwa_leave_item-subty"

    leavetypes = {
        'Vacation': '0100',
        'Medical': '0200',
        'Medical without MC': '0201',
        'Childcare with MC': '0430',
        'Childcare without MC': '0440',
        'Compassionate': '0810',
        'Parent Care': '0821',
    }

    if jsonstring[0]['leaveType'] in leavetypes:
        jsonstring[0][newxmltag] = leavetypes[jsonstring[0]['leaveType']]
        print(f'\t{jsonstring[0]["leaveType"]} leave detected')
    else:
        print("No matching leave type!")

    with open(filepath, 'w') as file:
        j.dump(jsonstring, file)

def process_dots(filepath, filename):
    dotsJson = {'Country1': '', 'Country2': '', 'Country3': '', 'Country4': '', 'Country5': '', 'Country6': '', 'Country7': '', 'Country8': '', 'Country9': '', 'Country10': '', 'Country11': '', 'Country12': '', 'Country13': '', 'Country14': '', 'Country15': '', 'departureDate': '', 'returnDate': '', 'departureTravelMode': '', 'returnTravelMode': '', 'travelReasons1': '', 'travelReasons2': '', 'travelReasons3': '', 'pf': ''}

    with open(filepath, "r") as file:
        jsonObj = j.load(file)

    #extract and assign countries
    countries = [''] * 15
    idx = 0
    for country in jsonObj[0]["country"]:
        print(country["Column1"])
        countries[idx] = country["Column1"]
        idx += 1

    for key, val in zip(dotsJson, countries):
        dotsJson[key] = val

    dotsJson["departureDate"] = jsonObj[0]["startDate"]
    dotsJson["returnDate"] = jsonObj[0]["endDate"]
    dotsJson["departureTravelMode"] = jsonObj[0]["departureTravelMode"]
    dotsJson["returnTravelMode"] = jsonObj[0]["returnTravelMode"]
    dotsJson["travelReasons1"] = jsonObj[0]["travelReasons"]

    # add pf from filename format: pf,email,datetime,(CloudHR)_TYPE.TXT
    dotsJson["pf"] = filename.split(",")[0]

    #add parent item field
    dotsJson = {"item": dotsJson}

    jsonstring = j.dumps(dotsJson)
    jsonstring = j.loads(jsonstring.replace("\"'", '"'))
    print(jsonstring)

    with open(filepath, 'w') as file:
        j.dump(jsonstring, file)

def process_for_calendar(filepath, filename):
    # retrieve email from filename format: pf,email,datetime,(CloudHR)_TYPE.TXT
    recipient = filename.split(",")[1]

    with open(filepath, "r") as file:
        filedata = file.read()
        jsonstring = j.loads(filedata)

    eventList = []
    for events in jsonstring:
        event = {}
        for key, val in events.items():
            event[key] = val

        eventList.append(event)

    for event in eventList:
        end = event["End"]
        start = event["Start"]
        subject = event["Subject"]
        venue = event["Venue"]

        #Retrieve start date & time
        startSplit = parser.parse(start).isoformat().split("T") #convert 1/4/2023 5:00 PM to 2023-04-01T17:00:00
        startDMY = startSplit[0].split("-")
        startHMS = startSplit[1].split(":")
        startY = startDMY[0]
        startM = startDMY[1]
        startD = startDMY[2]
        startHH = startHMS[0]
        startMM = startHMS[1]

        #Retrieve end date & time
        endSplit = parser.parse(end).isoformat().split("T")
        endDMY = endSplit[0].split("-")
        endHMS = endSplit[1].split(":")
        endY = endDMY[0]
        endM = endDMY[1]
        endD = endDMY[2]
        endHH = endHMS[0]
        endMM = endHMS[1]

        send_invite(recipient,endY,endM,endD,endHH,endMM,startY,startM,startD,startHH,startMM,subject,venue)

    shutil.move(filepath, os.path.join(processed_cal_path(), filename))

def process_for_sap(filepath, filename, type):
    with open(filepath, "r") as file:
        jsonObj = j.load(file)

    xmlstring = json2xml.Json2xml(jsonObj).to_xml()
    print(xmlstring)
    xmlfilename = filename.replace('.TXT', '.XML') #case sensitive!
    with open(xmlfilename, "w") as xmlfile:
        xmlfile.write(xmlstring)

    #modify the default type property value (dict)
    xmlTree = ET.parse(xmlfilename)
    rootElement = xmlTree.getroot()
    for element in rootElement.findall("item"):
        element.set('type', type)

    xmlTree.write(xmlfilename, encoding='UTF-8', xml_declaration=True)

    shutil.move(xmlfilename, os.path.join(sap_bucket(), xmlfilename))
    shutil.move(filepath, os.path.join(processed_sap_path(), filename))

# https://learnpython.com/blog/working-with-icalendar-with-python/
# https://www.baryudin.com/blog/entry/sending-outlook-appointments-python/
def send_invite(recipient, endy, endm, endd, endhh, endmm, starty, startm, startd, starthh, startmm, subject, venue):
    # Build the event
    cal = icalendar.Calendar()
    cal.add('prodid', '-//Unclass calendar event //example.com//')
    cal.add('version', '2.0')
    cal.add('method', "REQUEST")

    attendee = 'MAILTO:' + recipient
    event = icalendar.Event()
    attendee_email = vCalAddress(attendee)
    event.add('attendee', attendee_email, encode=0)
    organiser_email = vCalAddress('MAILTO:service@example.com')
    event.add('organizer', organiser_email)
    event.add('status', "confirmed")
    event.add('category', "Event")
    event.add('summary', subject)
    event.add('description', subject)
    event.add('location', venue)
    event.add('dtstart', dt.datetime(int(starty), int(startm), int(startd), int(starthh), int(startmm), 0, tzinfo=pytz.utc))
    event.add('dtend', dt.datetime(int(endy), int(endm), int(endd), int(endhh), int(endmm), 0, tzinfo=pytz.utc))
    event['uid'] = 1234
    event.add('priority', 5)
    event.add('sequence', 1)

    #event.add('attendee', attendee_email)
    #event.add('organizer', organiser_email)
    #event.add('description', "Meeting with vendor")
    #event.add('location',"Jurong East")
    #event.add('dtstart', dt.datetime(2022, 1, 25, 8, 0, 0, tzinfo=pytz.utc))
    #event.add('dtend', dt.datetime(2022, 1, 25, 10, 0, 0, tzinfo=pytz.utc))
    #event.add('dtstart', start)
    #event.add('dtend', tz.localize(dt.datetime.combine(self.date, dt.time(start_hour + 1, start_minute, 0))))
    #event.add('dtstamp', tz.localize(dt.datetime.combine(self.date, dt.time(6, 0, 0))))
    #event['uid'] = self.get_unique_id() # Generate some unique ID
    #event.add('created', tz.localize(dt.datetime.now()))
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
    msg["Subject"] = subject
    msg["From"] = organiser_email
    msg["To"] = attendee_email
    msg["Content-class"] = "urn:content-classes:calendarmessage"
    msg.attach(email.mime.text.MIMEText(subject))

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

    # Write calendar invite to disk
    directory = Path.cwd() / 'processed-cal'
    try:
        directory.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        print("Folder already exists")
    else:
        print("Folder was created")

    filename = endy + endm + endd + endhh + endmm + starty + startm + startd + starthh + startmm + ".ics"
    filepath = processed_cal_path() + "/" + filename

    f = open(os.path.join(directory, filename), 'wb')
    f.write(cal.to_ical())
    f.close()

    e = open(filepath, 'rb')
    ecal = icalendar.Calendar.from_ical(e.read())
    for component in ecal.walk():
        print(component.name)
        if component.name == "VEVENT":
            print(component.get("attendee"))
            print(component.get("description"))
            print(component.get("summary"))
            print(component.get("location"))
            print(component.decoded("dtstart"))
            print(component.decoded("dtend"))
    e.close()

# Main function
if __name__ == '__main__':

    for filename in os.listdir(inbox_path()):
        filePath = os.path.join(inbox_path(), filename)

        if os.path.isfile(filePath):
            print("Processing file", filePath)
            handle_xml_type(filePath, filename)
            print("Done processing", filePath, "\n")
