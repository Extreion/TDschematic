# Vectorworks AV Script
# Written by Samuel Sueur
# Made for Tillman Domotics LLP

# Imports
import math
import vs
import csv
from datetime import datetime


record_list = []
fields = []
excel_rows = []
current_record = ''
csv_header = []
use_old_labels = False

def scan_document(handle):
    global fields
    global current_record
    global use_old_labels
    cable_number = ''
    if use_old_labels: # AV-01 001
        ptr = vs.GetSymLoc(handle)
        av_cable_number = list(vs.GetSymName(handle))
        cable_number = vs.PickObject(ptr[0] + 8, ptr[1])
        if av_cable_number[2] == '0':
            av_cable_number[2] = '-'
        else:
            av_cable_number.insert(2, '-')
        av_cable_number = ''.join(av_cable_number)
        cable_number = '="' + vs.GetSymName(cable_number) + '"'
    else: #AV-1101
        av_cable_number = vs.GetText(handle)
    # Get coordinates of current handle
    cable_type = 'Record missing'
    cable_purpose = 'Record missing'
    cable_from = 'Record missing'
    cable_to = 'Record missing'
    cable_room = 'Record missing'
    cable_floor = 'Record missing'
    cable_info = vs.GetRecord(handle, 1)
    if cable_info is not None:
        record_name = vs.GetName(cable_info)
        cable_type = vs.GetRField(handle, record_name, 'Cable Type')
        cable_purpose = vs.GetRField(handle, record_name, 'Cable Purpose')
        cable_from = vs.GetRField(handle, record_name, 'From')
        cable_to = vs.GetRField(handle, record_name, 'To')
        cable_floor = vs.GetRField(handle, record_name, 'Floor')
        cable_room = vs.GetRField(handle, record_name, 'Room')
    if use_old_labels:
        excel_rows.append((cable_floor, cable_room, cable_type, av_cable_number,
                           cable_number, cable_purpose, cable_from, cable_to))
    else:
        excel_rows.append((cable_floor, cable_room, cable_type, av_cable_number,
                           cable_purpose, cable_from, cable_to))


# @@@@@@@@ MAIN @@@@@@@@

now = datetime.now()
vs.AlrtDialog("This script will produce a new cable schedule. The name of the file is:" + ' Cable_Schedule_' +
              now.strftime("%d%m%Y_%H%M") + '.csv')

use_old_labels = vs.YNDialog("Would you like to use the old AV labels?")
if use_old_labels:
    csv_header = ['Floor', 'Room', 'Cable Type', 'AV Number', 'Cable Number', 'Cable Purpose', 'From', 'To']
else:
    csv_header = ['Floor', 'Room', 'Cable Type', 'AV Number', 'Cable Purpose', 'From', 'To']

# Get all records (including plugin objects...)
number_of_records = vs.NumRecords('')



for i in range(1, number_of_records + 1):
    recordHandle = vs.GetRecord('', i)
    if not vs.IsPluginFormat(recordHandle):
        record_list.append(recordHandle)

# Loop through all records
for record in record_list:
    fields = []
    current_record = record.name
    # Get all record fields
    # for i in range(1, vs.NumFields(record) + 1):
        # fields.append(vs.GetFldName(record, i))
        # vs.AlrtDialog(','.join(fields))
    # Get all objects in the schematic with the same record attached
    vs.ForEachObject(scan_document, "R IN ['" + record.name + "']")

#Create csv file
csvfile = open('Cable_Schedule_' + now.strftime("%d%m%Y_%H%M") + '.csv', 'w', newline='')
excel_stream = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, dialect='excel')
excel_stream.writerow(csv_header)
# Sort rows
if use_old_labels:
    sorter = lambda x: (x[3], x[4])
else:
    sorter = lambda x: (x[3])
sorted_rows = sorted(excel_rows, key=sorter)
excel_stream.writerows(sorted_rows)
csvfile.close()
vs.AlrtDialog("Done!")
