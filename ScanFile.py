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
csv_header = ['Floor', 'Room', 'Cable Type', 'AV Number', 'Cable Number', 'Cable Purpose', 'From', 'To']


def scan_document(handle):
    global fields
    global current_record
    vs.AlrtDialog(vs.GetSymName(handle))
    # Get coordinates of current handle
    ptr = vs.GetSymLoc(handle)
    cable_number = vs.PickObject(ptr[0] + 8, ptr[1])
    vs.AlrtDialog(vs.GetSymName(cable_number))
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

    excel_rows.append([cable_floor, cable_room, cable_type, '',
                       '', cable_purpose, cable_from, cable_to])
    # excel_rows.append([])
    # for field in fields:
        # vs.AlrtDialog(vs.GetRField(handle, current_record, field))



# @@@@@@@@ MAIN @@@@@@@@

now = datetime.now()
vs.AlrtDialog("This script will produce a new cable schedule. The name of the file is:" + 'Cable_Schedule_' +
              now.strftime("%d%m%Y_%H%M") + '.csv')

# Get all records (including plugin objects...)
number_of_records = vs.NumRecords('')


for i in range(number_of_records):
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
excel_stream.writerows(excel_rows)
csvfile.close()
