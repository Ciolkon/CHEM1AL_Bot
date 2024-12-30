from datetime import datetime

sectionTimes = {
    "T 8-11": ("201", "202", "203"),
    "T 1-4": ("211", "212", "213", "214"),
    "W 1-4": ("311", "312", "313", "314"),
    "W 6-9": ("321", "322"),
    "Th 1-4": ("411", "412", "413", "414"),
    "F 1-4": ("511")}

gsiSections = {
    "Drew Salmon": ("201", "211"),
    "Andres Arraiz": ("202", "213"),
    "Diego Novoa": ("203", "312"),
    "Sunnie Kong": ("212", "412"),
    "Utkarsh Tiwari": ("313", "411"),
    "Aaron Guo": ("214", "322"),
    "Jack Feldner": ("314", "413"),
    "Anna Kurianowicz": ("321", "511"),
    "Henry Phan": ("311", "414")}

gsiEmails = {
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',
    "exampleName": 'example.email@mail.com',}

gsiTimes ={}
for gsi, sections in gsiSections.items():
    for section in sections:
        for time, section_list in sectionTimes.items():
            if section in section_list:
                gsiTimes.setdefault(gsi, []).append((time))

sectionSheets = {
    '201': 0,
    '202': 1,
    '203': 2,
    '211': 3,
    '212': 4,
    '213': 5,
    '214': 6,
    '311': 7,
    '312': 8,
    '313': 9,
    '314': 10,
    '321': 11,
    '322': 12,
    '411': 13,
    '412': 14,
    '413': 15,
    '414': 16,
    '511': 17,}

labDates = {
    'Lab1': datetime(2024, 1, 27),
    'Lab2': datetime(2024, 2, 3),
    'Lab3': datetime(2024, 2, 10),
    'Lab4': datetime(2024, 2, 17),
    'Lab5': datetime(2024, 2, 24),
    'Lab6': datetime(2024, 3, 2),
    'Lab7': datetime(2024, 3, 9),
    'Lab8': datetime(2024, 3, 16),
    'Lab9': datetime(2024, 3, 23),
    'Lab10': datetime(2024, 4, 6),
    'Lab11': datetime(2024, 4, 13),
    'LabFinal': datetime(2025, 1, 1)}   #This date is pushed back otherwise the attendance counter breaks at the end of the semester

labIndex = {                                                                    #Column index of Weekly Checklist for attendance functions
    'Lab1': 4,
    'Lab2': 7,
    'Lab3': 10,
    'Lab4': 13,
    'Lab5': 16,
    'Lab6': 19,
    'Lab7': 22,
    'Lab8': 25,
    'Lab9': 28,
    'Lab10': 31,
    'Lab11': 34,
    'LabFinal': 37}

subToLabel = {
    "Chem 1AL Lab Lecture Credit": "Lab Lecture",
    "Chem 1AL Extension Request Accepted": "Extension Requests (Accepted)",
    "Chem 1AL Extension Request Response": "Extension Requests (Denied)",
    "Chem 1AL Prelab Quiz": "Pre-Lab Quiz",
    "Chem 1AL Online Lab Instructions": "Online Lab",
    "Chem 1AL Rescheduled Lab": "Rescheduling Lab Session",
}

lab9StudentData = ['101', 
                   '201', 
                   '202', 
                   '203', 
                   '204', 
                   '211', 
                   '212', 
                   '213', 
                   '221', 
                   '311', 
                   '312', 
                   '313', 
                   '321', 
                   '322', 
                   '324', 
                   '401', 
                   '402', 
                   '403', 
                   '411', 
                   '412', 
                   '414', 
                   '511']