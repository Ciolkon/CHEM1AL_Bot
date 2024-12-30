#API Plugins

#Google Sheets API

import gspread
from oauth2client.service_account import ServiceAccountCredentials

sheets_scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
sheets_creds = ServiceAccountCredentials.from_json_keyfile_name("chem-1al-creds.json", sheets_scope)
client = gspread.authorize(sheets_creds)

#Gmail API

from googleApis import create_service
gmail_creds = 'gmail_creds.json'
api_name = 'gmail'
api_version = 'v1'
scope = ['https://mail.google.com/']
service = create_service(gmail_creds, api_name, api_version, scope)

#OpenAI API

import openai
from openai import OpenAI
openaiClient = OpenAI(api_key='example')
chem1alAssistantID = 'example'

#Sheet plugins

sheetContactForm = client.open("Contact Form Example")                             #Student contact sheet input
seqResponse = sheetContactForm.get_worksheet(1)                                    #Sequential fix
sheetAttendance = client.open("1AL Weekly Checklist Spring 2024")                  #Student Attendance info
print("Google sheets API implemented successfully")

#Gmail plugins

from email.message import EmailMessage
import base64
from email.mime.text import MIMEText

#Dictionary plugins

from CHEM1AL_Dicts import sectionTimes, gsiSections, gsiEmails, gsiTimes, sectionSheets, labDates, labIndex, subToLabel, lab9StudentData

#Other modules

import sys
import random
from datetime import datetime, timedelta
import re
import pytz

#Actual code

headGsiName = "Your Name"

timeIndex = 0
emailIndex = 1
fNameIndex = 2
lNameIndex = 3
issueIndex = 4
extReasonIndex = 5
extOtherIndex = 6
assignmentIndex = 7
dueDateIndex = 8
sectionIndex = 9
gsiIndex = 10
newTimeIndex = 11

def gsiFromSection(gsiSections, section):                                          #Parses the gsiSections dict for gsiName from section
    for gsi, sections in gsiSections.items():
        if section in sections:
            return gsi

def gsiEmailFromSection(gsiSections, gsiEmails, section):                          #Parses both the gsiSections and gsiEmails for gsiEmail from section
    for gsiName, sections in gsiSections.items():
        if section in sections:
            return gsiEmails.get(gsiName)

currentTime = datetime.now()
currentLab = None
for lab, date in labDates.items():                                                 #Determines current lab and assigns index number based on current time and labDates and labIndex dicts
    if currentTime <= date:
        currentLab = labIndex[lab]
        break

def readTimestamp():                                                               #Reads timestamp.txt file
    try:
        with open('timestamp.txt', 'r') as file:
            timestampStr = file.read()
            return datetime.strptime(timestampStr, "%Y-%m-%d %H:%M:%S")
    except (FileNotFoundError, Exception):
        return datetime(2023, 12, 18, 16, 50, 20)

def writeTimestamp(timestamp):                                                     #Writes the time of last student input to timestamp.txt file
    with open('timestamp.txt', 'w') as file:
        file.write(timestamp.strftime("%Y-%m-%d %H:%M:%S"))

def readContactForm(seqResponse):                                                  #Function that reads rows until last timestamp is encountered
    input = []
    lastCheck = readTimestamp()
    allRows = seqResponse.get_all_values()

    for rowValues in allRows[1:]:
        if any(rowValues):
            inputTime = datetime.strptime(rowValues[0], "%m/%d/%Y %H:%M:%S")
            if inputTime > lastCheck:
                input.append(rowValues)
        else:
            break
        
    if input:
        lastCheck = datetime.strptime(input[0][0], "%m/%d/%Y %H:%M:%S")
        writeTimestamp(lastCheck)
    return input

def gsiAnnoyer(studentInput):                                                      #Checks if GSI has completed attendance
    studentSection = studentInput[sectionIndex]
    sectionSheet = sheetAttendance.get_worksheet(sectionSheets[studentSection])
    lastLab = currentLab - 3
    if lastLab < 4:
        falsePercentage = 0
    else:
        colValues = sectionSheet.col_values(lastLab)
        labAttendance = colValues[2::1]
        total = len(labAttendance)
        falseCount = labAttendance.count('FALSE')
        falsePercentage = (falseCount/total) * 100
   
    if falsePercentage > 30:
        return False
    else:
        return True

def absenceCounter(studentInput):                                                  #Counts absences up to current lab
   studentEmail = studentInput[emailIndex]
   studentSection = studentInput[sectionIndex]
   sectionSheet = sheetAttendance.get_worksheet(sectionSheets[studentSection])
   columnIndex = 2
   columnValues = sectionSheet.col_values(columnIndex)
   rowIndex = columnValues.index(studentEmail) + 1
   rowCheckList = sectionSheet.row_values(rowIndex)
   lastLab = currentLab - 3
   absences = rowCheckList[3:lastLab:3]
   #print(f'Online lab test --- lab check: {lastLab}')                              #Helpful print statement for debugging
   if lastLab < 4:
       absenceTotal = 0
   else:
       absenceTotal = sum([int(boolean == 'FALSE') for boolean in absences])
   
   return absenceTotal

def courseAbsenceCounter(currentSheet):                                            #Counts absences for the entire course and prints students with >= 3 absences
    lastLab = currentLab - 3
    studentNameIndex = 0
    studentEmailIndex = 1
    
    sheetData = currentSheet.get_all_values()
    for row in sheetData[2:]:
        absences = row[3:lastLab:3]
        if lastLab < 4:
            absenceTotal = 0
        else:
            absenceTotal = sum([int(boolean == 'FALSE') for boolean in absences])
        if absenceTotal >= 3:
            print(f'Attendance issue: {row[studentNameIndex]}, {row[studentEmailIndex]}, {currentSheet}, Absence Total = {absenceTotal}')
    return

def labReportCounter(currentSheet):                                                #Counts report submissions for the entire course and prints students with >= 3 missing
    lastLab = currentLab - 2
    studentNameIndex = 0
    studentEmailIndex = 1
    
    sheetData = currentSheet.get_all_values()
    for row in sheetData[2:]:
        reports = row[4:lastLab:3]
        arguementation = row[36]
        if lastLab < 7:
            missingReportsTotal = 0
        else:
            missingReportsTotal = sum([int(boolean == 'FALSE') for boolean in reports]) + int(arguementation == 'FALSE')        #<==  note out before Arg. deadline
        if missingReportsTotal >= 3:
            print(f'Lab Report Issue: {row[studentNameIndex]}, {row[studentEmailIndex]}, {currentSheet}, Missing Report Total = {missingReportsTotal} ')
    return

def rosterCheck(studentInput):                                                     #Checks if student is actually in the section they say they are
    studentSection = studentInput[sectionIndex]
    studentEmail = studentInput[emailIndex]
    sectionSheet = sheetAttendance.get_worksheet(sectionSheets[studentSection])
    columnIndex = 2
    studentEmails = sectionSheet.col_values(columnIndex)
    if studentEmail in studentEmails:
        return True
    else:
        return False

def draftEmail(emailContent, recipient, ccRecipient1, ccRecipient2, subject):      #General draft writing method for gmail API

   message = EmailMessage()
   message.set_content(emailContent)
   message["To"] = recipient
   message["Cc"] = ccRecipient1, ccRecipient2
   message["From"] = "chem1al.sp24@gmail.com"
   message["Subject"] = subject

   encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
   createMessage = {"message": {"raw": encoded_message}}

   draft = service.users().drafts().create(userId="me", body=createMessage).execute()
   return draft

def labLectureResponse(studentInput):                                              #Makes draft with copy paste of lecture instructions
   template = """Hello {studentName},

There are 2 participation points for attending each lecture, please follow the instructions below to receive credit.

1. Attending lecture and answering ALL iClicker questions. If you have technical difficulties for any reason or are absent from lecture you move on to option 2
2. Completing the makeup form on Bcourses. You can find this in the lecture module for this week. The deadline for completing the makeup form is Monday at 11:59 PM. Note: If you are completing a makeup form your grade will not be automatically updated. I have to do this manually so there's a delay. Lecture grades will be updated by the end of the day every Tuesday.

If you're trying to receive credit for a past week's lecture and the participation form is no longer open then you will be unable to make up credit.

    Best,
    {headGsiName}
   """
   global headGsiName
   studentName = studentInput[fNameIndex]
   emailContent = template.format(studentName=studentName, headGsiName=headGsiName)

   draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Lab Lecture Credit")
   return draft

def extensionAccepted(studentInput):                                               #Makes draft accepting the extention request
   template = """Hello {studentName},

I'll extend your deadline until Sunday at 11:59 pm for the {labSession} {assignment}. That is a hard deadline though, we release grades every Monday so I can't extend past that.

    Best,
    {headGsiName}
   """
   global headGsiName
   studentName = studentInput[fNameIndex]
   labSession = studentInput[assignmentIndex]
   assignment = studentInput[issueIndex]
   emailContent = template.format(studentName=studentName, labSession=labSession, assignment=assignment, headGsiName=headGsiName)

   draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Extension Request Accepted")
   return draft

def extensionDeniedExcuse(studentInput):                                           #Makes draft denying the extention request because the reason is invalid
   template = """Hello {studentName},
This is in response to your extension request for {assignment}.
   
I'm sorry, but we can't extend your deadline for this assignment. Please review the policy in the syllabus, for most assignments the penalty is very modest (1 point per day).
Please try to submit your best work as soon as you can.

    Best,
    {headGsiName}
   """
   global headGsiName
   studentName = studentInput[fNameIndex]
   assignment = studentInput[assignmentIndex]
   emailContent = template.format(studentName=studentName, assignment=assignment, headGsiName=headGsiName)

   draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Extension Request Response")
   return draft

def extensionDeniedLate(studentInput):                                             #Makes draft denying the extention request because they submitted their request too late
   template = """Hello {studentName},
This is in response to your extension request for {assignment}.
   
I'm sorry, while we do offer accommodations we can not when a student reaches out after or within hours of the assignment deadline. Please check the course late policy and try to submit your best work as soon as you can. For most assignments the penalty is modest. (1 point per day)

    Best,
    {headGsiName}
   """
   global headGsiName
   studentName = studentInput[fNameIndex]
   assignment = studentInput[assignmentIndex]
   emailContent = template.format(studentName=studentName, assignment=assignment, headGsiName=headGsiName)

   draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Extension Request Response")
   return draft

def extensionAttnRequired(studentInput):                                           #Makes draft for special needs extention requests
   template = """{studentName} {studentEmail} Needs special attention on an extension request ({assignment}) for the following reason:
   {studentReason}, {studentCase}
    ---
   """
   global headGsiName
   studentName = studentInput[fNameIndex]
   studentEmail = studentInput[emailIndex]
   studentCase = studentInput[extOtherIndex]
   studentReason = studentInput[extReasonIndex]
   assignment = studentInput[assignmentIndex]
   emailContent = template.format(studentName=studentName, studentEmail=studentEmail, studentCase=studentCase, studentReason=studentReason, assignment=assignment, headGsiName=headGsiName)

   draft = draftEmail(emailContent, studentEmail, '', '', "Chem 1AL Extension Request - Attention Required")
   return draft

def preLab(studentInput):                                                          #Makes draft about reopening pre-lab quiz
   template = """Hello {studentName},

I've reopened the {labSession} quiz for you. Please be sure to complete it before your lab session. Feel free to attend office hours or reach out to your GSI if you need any help.

If you're receiving this message after your lab has started then you will have to fill out the contact form again to request a make up lab. You can either do the online lab or reschedule in person.

    Best,
    {headGsiName}
   """
   global headGsiName
   studentName = studentInput[fNameIndex]
   labSession = studentInput[assignmentIndex]
   emailContent = template.format(studentName=studentName, labSession=labSession, headGsiName=headGsiName)

   draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Prelab Quiz Reopened")
   return draft

def onlineLab(studentInput):                                                       #Makes draft with online lab instructions depending on student attendance
   template1 = """Hey {gsiFirstName},

{studentName} is requesting to do the online lab for {labSession} but I can't confirm their attendance because your checklist for section {studentSection} isn't complete. Could you please complete that so I can respond to them?

    Thanks,
    {headGsiName}
   """
   
   template2 = """Hello {studentName},

Make sure you understand the attendance policy in the syllabus. Each student can only do two online labs. I've copied the instructions to the online lab below and CC'd your GSI so they're aware.

Instructions:
Complete the online lab. There is a link on the Bcourses homepage.
    
As you work through the online lab you should fill out your lab notebook and lab report pages as if you were doing the lab in person. Base your observations on the videos included in the online lab. You can find guidelines for notebook pages on page 13 of your lab manual.
    
Complete your lab report and submit to Gradescope, the due date is the night before your next lab session so this is dependent on the section you are assigned to. When asked to write down the names of your lab partners please note that you did the experiment online instead. This will help us when we're grading.
Please reach out to your GSI if you have any questions or need help. 

    Best,
    {headGsiName}
   """
   template3 = """Hello {studentName},

According to our records you have {absenceTotal} absences already this semester. This means we can no longer offer you online lab accommodations. Please follow up as soon as you receive this message to reschedule in person.

    Best,
    {headGsiName}
   """
   template4 = """Hello {studentName} 
   
   I can't address your issue because you've listed section {studentSection} with {gsiName} as your assigned section which is incorrect according to our records. Please double check your section assignment and resubmit the form as soon as possible to have your issue addressed.
   
   Best,
   {headGsiName}"""
       
   global headGsiName
   studentName = studentInput[fNameIndex]
   gsiName = gsiFromSection(gsiSections, studentInput[sectionIndex])
   gsiNameSplit = gsiName.split()
   gsiFirstName = gsiNameSplit[0]
   studentSection = studentInput[sectionIndex]
   labSession = studentInput[assignmentIndex]                            #Check here
   correctSection = rosterCheck(studentInput)
   #studentData = random.choice(lab9StudentData)

   if correctSection == True:
        gsiAttendance = gsiAnnoyer(studentInput)
        absenceTotal = absenceCounter(studentInput)
        print(f'Attendance Check: {studentInput[fNameIndex]} {studentInput[lNameIndex]} {studentInput[sectionIndex]} --- Absence Total: {absenceTotal}')
        if gsiAttendance == False:
            emailContent = template1.format(gsiFirstName=gsiFirstName, studentSection=studentSection, studentName=studentName, labSession=labSession, headGsiName=headGsiName)

            draft = draftEmail(emailContent, gsiEmailFromSection(gsiSections, gsiEmails, [studentInput[sectionIndex]]), studentInput[emailIndex], '', "Attendance Checklist Issue") 
        elif absenceTotal >= 2:
            emailContent = template3.format(studentName=studentName, absenceTotal=absenceTotal, headGsiName=headGsiName)

            draft = draftEmail(emailContent, studentInput[emailIndex], gsiEmailFromSection(gsiSections, gsiEmails, studentInput[sectionIndex]), '', "Chem 1AL Attendance Issue - Attention Required")
        else:
            emailContent = template2.format(studentName=studentName, absenceTotal=absenceTotal, headGsiName=headGsiName)

            draft = draftEmail(emailContent, studentInput[emailIndex], gsiEmailFromSection(gsiSections, gsiEmails, studentInput[sectionIndex]), '', f"Chem 1AL Online Lab Instructions - {labSession}")
   else:
        emailContent = template4.format(studentName=studentName, studentSection=studentSection, gsiName=gsiName, headGsiName=headGsiName)

        draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Attention Required: Incorrect Section")    
   return draft

def labReschedule(studentInput):                                                   #Lab rescheduling function
   
   global headGsiName

   correctSection = rosterCheck(studentInput)
   studentSection = studentInput[sectionIndex]
   studentFName = studentInput[fNameIndex]                                         #Randomly assigns a new time based on student options
   studentLName = studentInput[lNameIndex]
   timeOptions = studentInput[newTimeIndex]
   timeOptionsSplit = [option.strip() for option in timeOptions.split(',')]

   newTime = random.choice(timeOptionsSplit)

   availableGSIs = [gsiName for gsiName, timeList, in gsiTimes.items() if newTime in timeList]  #Randomly picks new GSI in selected timeslot
   newGsi = random.choice(availableGSIs)

   gsiNameSplit = newGsi.split()
   newGsiFirstName = gsiNameSplit[0]

   newGsiSections = gsiSections.get(newGsi, ())
   for section in newGsiSections:
        if section in sectionTimes.get(newTime, ()):
            newSection = section

   template1 = """Hello {studentName},                                              

I'll have you join {newGsi}'s section {newSection}, {newTime}. You can find the room number on Bcourses. I've CC'd both GSI's here so that they're aware.

{newGsiFirstName}, can you fill out the attendance for {studentFName} {studentLName} in the section {studentSection} spreadsheet with a note that they rescheduled?

    Best,
    {headGsiName}
   """
   template2 = """Hello {studentName},

I'll just have you join {newGsi}'s other section {newSection}, {newTime}. I've CC'd them so they're aware. You can find the room number on Bcourses.

    Best,
    {headGsiName}
   """
   
   template3 = """Hello {studentName}, 
   
   I can't address your issue because you've listed section {studentSection} as your assigned section which is incorrect according to our records. Please double check your section assignment and resubmit the form as soon as possible to have your issue addressed.
   
   Best,
   {headGsiName}"""

   studentName = studentInput[fNameIndex]
   if correctSection == False:
       emailContent = template3.format(studentName=studentName, studentSection=studentSection, headGsiName=headGsiName)
       
       draft = draftEmail(emailContent, studentInput[emailIndex], '', '', "Chem 1AL Attention Required: Incorrect Section")
   elif newGsi == gsiFromSection(gsiSections, studentInput[sectionIndex]):
        emailContent = template2.format(studentName=studentName, newGsi=newGsi, newSection=newSection, studentSection=studentSection, newTime=newTime, newGsiFirstName=newGsiFirstName, headGsiName=headGsiName)        

        draft = draftEmail(emailContent, studentInput[emailIndex], gsiEmailFromSection(gsiSections, gsiEmails, studentInput[sectionIndex]), '', f"Chem 1AL Rescheduled Lab --- {studentInput[assignmentIndex]}")
   else:
    emailContent = template1.format(studentName=studentName, newGsi=newGsi, newSection=newSection, studentSection=studentSection, newTime=newTime, newGsiFirstName=newGsiFirstName, studentFName=studentFName, studentLName=studentLName, headGsiName=headGsiName)

    draft = draftEmail(emailContent, studentInput[emailIndex], gsiEmailFromSection(gsiSections, gsiEmails, studentInput[sectionIndex]), gsiEmails[newGsi], f"Chem 1AL Rescheduled Lab --- {studentInput[assignmentIndex]}")
    
   return draft

def empathy(studentInput):                                                         #Catch all for all other form submissions
   template = """{studentName} {studentEmail} Needs special attention for the following reason:
    {studentCase}
    ---
   """
   studentName = studentInput[fNameIndex] + " " + studentInput[lNameIndex]
   studentEmail = studentInput[emailIndex]
   studentCase = studentInput[extOtherIndex]
   emailContent = template.format(studentName=studentName, studentEmail=studentEmail, studentCase=studentCase)

   draft = draftEmail(emailContent, studentEmail, '', '', "Student Needs Special Attention")
   return draft

def mimeParsing(part):                                                             #Recursive function for parsing gmail message data
    if part['mimeType'] == 'text/plain' or part['mimeType'] == 'text/html':
        bodyData = part['body']['data']
        body = base64.urlsafe_b64decode(bodyData).decode('utf-8')
        return body
    elif 'parts' in part:
        for subpart in part['parts']:
            body = mimeParsing(subpart)
            if body:
                return body
    return None

def getFullEmail(service, user_id, msg_id):                                        #Parses and retrieves email body
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id, format='full').execute()
        part = message['payload']
        body = mimeParsing(part)

        return body
    except Exception as e:
        print(f'An error occured: {e}')

def getNewEmails(service, user_id='me'):                                           #Creates list of dictionaries of new emails with content
    
    try:
        response = service.users().messages().list(userId=user_id, labelIds=['INBOX'], q='is:unread').execute()
        messages = response.get('messages', [])

        emails = []
        for message in messages:
            msg = service.users().messages().get(userId=user_id, id=message['id'], format='metadata').execute()
            
            headers = msg['payload']['headers']
            subject = next(header['value'] for header in headers if header['name'] == 'Subject')
            sender = next(header['value'] for header in headers if header['name'] == 'From')

            full_body = getFullEmail(service, user_id, message['id'])
            emails.append({'id':message['id'], 'body': full_body, 'subject': subject, 'sender': sender})
        return emails
    
    except Exception as e:
        print(f'An error occured: {e}')
        return None

def sendChatGPT(emailBody, assistantID):                                           #Sends emails to chatGPT assistant
    
    thread = openaiClient.beta.threads.create()
    
    message = openaiClient.beta.threads.messages.create(
        thread_id = thread.id,
        role = "user",
        content = emailBody
    )
  
    run = openaiClient.beta.threads.runs.create(
        thread_id = thread.id,
        assistant_id = assistantID
    )

    while run.status not in ["completed", "failed"]:
        run = openaiClient.beta.threads.runs.retrieve(
            thread_id = thread.id,
            run_id = run.id
        )

    if run.status == "completed":
        messages = openaiClient.beta.threads.messages.list(
            thread_id = thread.id
        )
        for message in messages:
            if message.role == 'assistant':
                for content in message.content:
                    rawResponse = content.text.value.replace('\n', ' ')
                    break
                response = re.sub(r"【.*?】", "", rawResponse)
        print('response received')
        return response

def chatgptDraft(service, user_id, recipient, subject, response, thread_id):       #Creates draft with chatGPT assistant response
    template = """Hello, 
    {chatgptResponse}
    Best,
    {headGsiName}"""
    
    chatgptResponse = response
    emailContent = template.format(chatgptResponse=chatgptResponse)
    message = MIMEText(emailContent)
    message['to'] = recipient
    message['subject'] = subject

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    draft_body = {'message': {'raw': raw_message}}

    if thread_id:
        draft_body['message']['threadId'] = thread_id

    draft = service.users().drafts().create(userId=user_id, body=draft_body).execute()
    print(f"Stochastic parrot response: {draft['id']}")  
    return draft

def getSentEmails(service, user_id='me'):                                          #Gets sent emails within the last 24hrs as msgID: subject list of dicts
    within24Hrs = datetime.now(pytz.utc) - timedelta(days=1)
    queryDate = within24Hrs.strftime('%Y/%m/%d')

    try:
        response = service.users().messages().list(userId=user_id, labelIds=['SENT'], q='after:' + queryDate).execute()
        messages = response.get('messages', [])

        sentEmails = []
        for message in messages:
            msg = service.users().messages().get(userId=user_id, id=message['id'], format='metadata').execute()
            
            headers = msg['payload']['headers']
            subject = next(header['value'] for header in headers if header['name'] == 'Subject')

            sentEmails.append({'id':message['id'], 'subject': subject})
        return sentEmails

    except Exception as e:
        print(f'An error occured: {e}')
        return None

def getLabelID(service, labelName, user_id='me'):                                  #Gets label IDs from gmail account
    try:
        response = service.users().labels().list(userId=user_id).execute()
        labels = response.get('labels', [])

        for label in labels:
            if label["name"] == labelName:
                return label['id']
        return None
    except Exception as e:
        print(f"An error occurred retrieving labels: {e}")

def applyLabel(service, user_id, message_id, label_id):                            #Applys label IDs to messages depending on subject line
    try:
        body = {'addLabelIds': [label_id], 'removeLabelIds': []}
        service.users().messages().modify(userId=user_id, id=message_id, body=body).execute()
    except Exception as e:
        print(f"An error occurred while applying label: {e}")

def main():                                                                     
    command = " ".join(sys.argv[1:])
    if command != 'You write my emails' and command != "You are a chatGPT wrapper" and command != "You sort my emails" and command != "You check attendance" and command != 'You count lab reports':
        print('What is my purpose?')
        sys.exit(1)

    if command == 'You write my emails':                                                #Answers contact form inputs
        studentResponse = readContactForm(seqResponse)
        for row in studentResponse:                                                     #Splits outcomes based on student input
            if row[issueIndex] == "Missed lab lecture or missing lab lecture credit":
                print("Lab lecture")
                try:
                    labLectureResponse(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "Lab report sheet/ notebook" and row[extReasonIndex] == "Illness" and (datetime.strptime(row[dueDateIndex] + " 23:59", "%m/%d/%Y %H:%M") > datetime.strptime(row[timeIndex], "%m/%d/%Y %H:%M:%S") and datetime.strptime(row[dueDateIndex] + " 23:59", "%m/%d/%Y %H:%M") > currentTime):
                print(f"Extention request accepted {row[fNameIndex]} {row[lNameIndex]} {row[assignmentIndex]} --- *Attention Required*")
                try:
                    extensionAccepted(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "Lab report sheet/ notebook" and row[extReasonIndex] == "Illness" and (datetime.strptime(row[dueDateIndex] + " 23:59", "%m/%d/%Y %H:%M") < datetime.strptime(row[timeIndex], "%m/%d/%Y %H:%M:%S") or datetime.strptime(row[dueDateIndex] + " 23:59", "%m/%d/%Y %H:%M") < currentTime):
                print('Extention request denied, late')
                try:
                    extensionDeniedLate(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "Lab report sheet/ notebook" and (row[extReasonIndex] == "Overwhelmed with course work"):
                print('Extention request denied, bad excuse')
                try:
                    extensionDeniedExcuse(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "Lab report sheet/ notebook" and (row[extReasonIndex] == "Other" or row[extReasonIndex] == 'N/A' or row[extReasonIndex] == 'DSP Accommodation'):
                print(f"Extention request attention required --- {row[assignmentIndex]}")
                try:
                    extensionAttnRequired(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "Pre-lab quiz":
                print(f"Pre-lab quiz {row[fNameIndex]} {row[lNameIndex]} {row[sectionIndex]} {row[assignmentIndex]} --- *Attention Required*")
                try:
                    preLab(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "In-person lab" and ("I am requesting to do the online lab (Max 2)" in row[newTimeIndex]):
                print('Online lab')
                try:
                    onlineLab(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "In-person lab" and ("I am requesting to do the online lab (Max 2)" not in row[newTimeIndex]):
                print('Lab reschedule')
                try:
                    labReschedule(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
            if row[issueIndex] == "Other":
                print('Special case')
                try:
                    empathy(row)
                except Exception as e:
                    print(f'An error occured: {e}, {row[fNameIndex]}, {row[lNameIndex]}, {row[emailIndex]}, {row[issueIndex]}')
        print('Oh my god...')
        sys.exit(0)

    if command == "You are a chatGPT wrapper":                                          #Call CHEM1AL chatGPT assistant to answer student emails
        newEmails = getNewEmails(service)
        for email in newEmails:
            emailBody = email['body']
            sender = email['sender']
            subject = "Re: " + email['subject']
            try:
                response  = sendChatGPT(emailBody, chem1alAssistantID)
                chatgptDraft(service, 'me', sender, subject, response, email['id'])
            except Exception as e:
                print(f"An error occured: {e}, {sender}, {subject}")
        print('Oh my god...')
        sys.exit(0)

    if command == "You sort my emails":                                                 #Sorts sent emails into the correct folders
        sentEmails = getSentEmails(service, user_id='me')
        
        if sentEmails:
            labelNametoID = {name: getLabelID(service, name) for name in subToLabel.values()}
            
            for emailDict in sentEmails:
                subject = emailDict.get('subject', '')
                messageID = emailDict.get('id', '')

                labelApplied = False

                if subject in subToLabel:
                    labelName = subToLabel[subject]
                    labelID = labelNametoID.get(labelName)
                    if labelID:
                        applyLabel(service, 'me', messageID, labelID)
                        print(f"Label '{labelName}' applied to email with subject: '{subject}'")
                        labelApplied = True

                if not labelApplied:
                    for key, labelName in subToLabel.items():
                        if subject.startswith(key):
                            labelID = labelNametoID.get(labelName)
                            if labelID:
                                applyLabel(service, 'me', messageID, labelID)
                                print(f"Label '{labelName}' applied to email with subject: '{subject}'")
                            break
        else:
            print('No new sent emails')
        print('Oh my god...')
        sys.exit(0)

    if command == "You check attendance":                                               #Check student attendance and finds students with >= 3 absences
        sectionSheets = sheetAttendance.worksheets()
        
        for currentSheet in sectionSheets:
            courseAbsenceCounter(currentSheet)
        print("Oh my god")
        sys.exit(0)

    if command == "You count lab reports":                                              #Check report submissions and finds students with >= 3 missing
        sectionSheets = sheetAttendance.worksheets()
        
        for currentSheet in sectionSheets:
            labReportCounter(currentSheet)
        print("Oh my god")
        sys.exit(0)        

if __name__ == '__main__':
    main()