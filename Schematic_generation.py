# Vectorworks AV Script
# Written by Samuel Sueur
# Made for Tillman Domotics LLP

# Imports
import math
import vs
import csv
from datetime import datetime

# Variable used to offset the starting column of the schematic in case of new columns in the excel formatting


offset_array_start = 0
offset_value = 2  # for room properties
offset_equipment_name = 1  # for equipment name
offset_equipment_quantity = 2  # for equipment quantity


# Class definition
class Schematic:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.local_number = 1
        self.tv_number = 1
        self.de_number = 1
        self.cctv_number = 1


class Page:
    def __init__(self, number, x, y):
        self.number = number
        self.x = x
        self.y = y
        self.width = 288
        self.height = 200
        self.drawing_area_width = 220
        self.drawing_area_height = 192
        self.drawing_area_x = 60
        self.y_pointer = -4


class Room:
    def __init__(self, x, y):
        self.width = 164
        self.height = 0
        self.av_number = 1
        self.cable_number = 0
        self.x = 0
        self.y = 0
        self.room = ""
        self.floor = ""
        self.labels_created = 0


schematic = Schematic(-1008, 500)
page = Page(0, schematic.x, schematic.y)
room_info = Room(page.x + page.drawing_area_x, page.y + page.y_pointer)  # create a new room

# Text
arial = vs.GetFontID("Arial")
text_size = 8

# Colours
colour_white = (65535, 65535, 65535)
colour_black = (0, 0, 0)
colour_driver_blue = (26214, 21845, 52428)
colour_wattage_red = (48059, 0, 0)
colour_mains_grey = (34952, 34952, 34952)
colour_alpha_brown = (19789, 17476, 6425)

# Room label variables
rl_width = 36
rl_height = 4
rl_diamX = 4
rl_diamY = 4

# Excel file
excel_rows = []
excel_stream = 0


def replace_tv_numbers(handle):
    global use_tv_numbers
    global schematic
    if use_tv_numbers:
        if schematic.tv_number >= 100:
            tv_name = "TV" + str(schematic.tv_number)
        else:
            if 100 > schematic.tv_number >= 10:
                tv_name = "TV0" + str(schematic.tv_number)
            else:
                tv_name = "TV00" + str(schematic.tv_number)
        vs.SetHDef(handle, vs.GetObject(tv_name))
        schematic.tv_number += 1
    return


def replace_local_numbers(handle):
    global use_local_numbers
    global schematic
    if use_local_numbers:
        if schematic.local_number >= 100:
            local_name = "Local" + str(schematic.local_number)
        else:
            if 100 > schematic.local_number >= 10:
                local_name = "Local0" + str(schematic.local_number)
            else:
                local_name = "Local00" + str(schematic.local_number)
        vs.SetHDef(handle, vs.GetObject(local_name))
        schematic.local_number += 1
    return


def replace_av_numbers(handle):
    global room_info
    global use_cable_numbers, use_av_numbers
    global excel_rows
    av_number_string = ""
    cable_number_string = ""
    if use_av_numbers:
        if room_info.av_number >= 100:
            av_name = "AV" + str(room_info.av_number)
            av_number_string = "AV-" + str(room_info.av_number)
        else:
            if 100 > room_info.av_number >= 10:
                av_name = "AV0" + str(room_info.av_number)
                av_number_string = "AV-" + str(room_info.av_number)
            else:
                av_name = "AV00" + str(room_info.av_number)
                av_number_string = "AV-0" + str(room_info.av_number)
        if 10 <= room_info.cable_number < 100:
                cable_number_string = '="0' + str(room_info.cable_number) + '"'
        else:
            if room_info.cable_number < 10:
                cable_number_string = '="00' + str(room_info.cable_number) + '"'
        if produce_excel_schedule:
            cable_type = 'Record missing'
            cable_purpose = 'Record missing'
            cable_from = 'Record missing'
            cable_to = 'Record missing'
            cable_info = vs.GetRecord(handle, 1)
            if cable_info is not None:
                record_name = vs.GetName(cable_info)
                cable_type = vs.GetRField(handle, record_name, 'Cable Type')
                cable_purpose = vs.GetRField(handle, record_name, 'Cable Purpose')
                cable_from = vs.GetRField(handle, record_name, 'From')
                cable_to = vs.GetRField(handle, record_name, 'To')
            excel_rows.append([room_info.floor, room_info.room, cable_type, av_number_string,
                               str(cable_number_string), cable_purpose, cable_from, cable_to])
        vs.SetHDef(handle, vs.GetObject(av_name))
    if use_cable_numbers:
        if room_info.cable_number >= 100:
            cable_name = str(room_info.cable_number)
        else:
            if 100 > room_info.cable_number >= 10:
                cable_name = "0" + str(room_info.cable_number)
            else:
                cable_name = "00" + str(room_info.cable_number)
        label_pos_tuple = vs.GetSymLoc(handle)
        label_pos_list = list(label_pos_tuple)
        label_pos_list[0] += 8
        label_pos_tuple = tuple(label_pos_list)
        vs.Symbol(cable_name, label_pos_tuple, 0)
        room_info.cable_number += 1
    return


def replace_cctv_numbers(handle):
    global schematic
    global use_cctv_numbers
    global excel_rows
    if use_cctv_numbers:
        if produce_excel_schedule:
            cable_type = 'Record missing'
            cable_purpose = 'Record missing'
            cable_from = 'Record missing'
            cable_to = 'Record missing'
            cable_info = vs.GetRecord(handle, 1)
            if cable_info is not None:
                record_name = vs.GetName(cable_info)
                cable_type = vs.GetRField(handle, record_name, 'Cable Type')
                cable_purpose = vs.GetRField(handle, record_name, 'Cable Purpose')
                cable_from = vs.GetRField(handle, record_name, 'From')
                cable_to = vs.GetRField(handle, record_name, 'To')
            excel_rows.append([room_info.floor, room_info.room, cable_type, "CCTV",
                               schematic.cctv_number, cable_purpose, cable_from, cable_to])
        vs.SetHDef(handle, vs.GetObject("CCTV"))
    if use_cable_numbers:
        if schematic.cctv_number >= 100:
            cable_name = str(schematic.cctv_number)
        else:
            if 100 > schematic.cctv_number >= 10:
                cable_name = '0' + str(schematic.cctv_number)
            else:
                cable_name = '00' + str(schematic.cctv_number)
        label_pos_tuple = vs.GetSymLoc(handle)
        label_pos_list = list(label_pos_tuple)
        label_pos_list[0] += 8
        label_pos_tuple = tuple(label_pos_list)
        vs.Symbol(cable_name, label_pos_tuple, 0)
        schematic.cctv_number += 1
    return


def replace_de_numbers(handle):
    global schematic
    global use_de_numbers
    global excel_rows
    if use_de_numbers:
        if produce_excel_schedule:
            cable_type = 'Record missing'
            cable_purpose = 'Record missing'
            cable_from = 'Record missing'
            cable_to = 'Record missing'
            cable_info = vs.GetRecord(handle, 1)
            if cable_info is not None:
                record_name = vs.GetName(cable_info)
                cable_type = vs.GetRField(handle, record_name, 'Cable Type')
                cable_purpose = vs.GetRField(handle, record_name, 'Cable Purpose')
                cable_from = vs.GetRField(handle, record_name, 'From')
                cable_to = vs.GetRField(handle, record_name, 'To')
            excel_rows.append([room_info.floor, room_info.room, cable_type, "DE",
                               schematic.de_number, cable_purpose, cable_from, cable_to])
        vs.SetHDef(handle, vs.GetObject("Door Ent"))
    if use_cable_numbers:
        if schematic.de_number >= 100:
            cable_name = str(schematic.de_number)
        else:
            if 100 > schematic.cctv_number >= 10:
                cable_name = "0" + str(schematic.de_number)
            else:
                cable_name = "00" + str(schematic.de_number)
        label_pos_tuple = vs.GetSymLoc(handle)
        label_pos_list = list(label_pos_tuple)
        label_pos_list[0] += 8
        label_pos_tuple = tuple(label_pos_list)
        vs.Symbol(cable_name, label_pos_tuple, 0)
        schematic.de_number += 1
    return


def create_room_label(value, x_l, y_l):
    vs.BeginGroup()
    vs.FillPat(1)
    vs.FillBack(colour_white)
    vs.PenBack(colour_white)
    vs.PenFore(colour_black)
    # Rounded rectangle
    vs.RRect((x_l, y_l), (x_l + rl_width, y_l - rl_height), rl_diamX, rl_diamY)
    # Line weight
    vs.SetLW(vs.LNewObj(), 10)
    # Create text
    vs.FillPat(0)
    vs.TextOrigin((x_l + rl_width / 2, y_l - 0.6))
    vs.TextFont(arial)
    vs.TextSize(text_size)
    vs.CreateText(value)
    vs.SetTextJust(vs.LNewObj(), 2)
    vs.EndGroup()


def def_room_start(current_row, schematic_tab):
    # global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y
    # global room_name, room_floor, room_labels_created, room_floor_x, room_floor_y, room_name_x, room_name_y
    # global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    global room_info, page
    room_info.x = page.x + page.drawing_area_x
    room_info.y = page.y + page.y_pointer
    room_info.height = 0
    while schematic_tab[current_row][offset_array_start] != "ROOM STOP":
        property = schematic_tab[current_row][offset_array_start]
        value = schematic_tab[current_row][offset_array_start + offset_value]
        if property == "ROOM FLOOR":
            room_info.floor = value
        if property == "ROOM NAME":
            room_info.room = value
        current_row += 1
    return current_row


def goto_next_page():
    global room_info, page
    if room_info.labels_created:
        # draw rectangle around room
        vs.FillPat(0)
        vs.PenBack(colour_black)
        vs.PenFore(colour_black)
        vs.Rect((room_info.x, room_info.y), (room_info.x + room_info.width, room_info.y - room_info.height))
        room_info.labels_created = 0
    page.number += 1
    if page.number % 7 == 0:
        page.x = schematic.x
    else:
        page.x = page.x + page.width
    page.y = schematic.y - math.floor((page.number / 7)) * page.height
    room_info.x = page.x + page.drawing_area_x
    room_info.y = page.y - 4
    page.y_pointer = -4
    room_info.height = 0
    vs.Symbol("Central System Rack Main", (page.x + 4, page.y - 4), 0)
    return


def def_equipment_start(current_row, schematic_tab):
    global room_info, page
    room_info.labels_created = 0
    current_row += 1
    while schematic_tab[current_row][offset_array_start] != "ROOM EQUIPMENT END":
        equipment = schematic_tab[current_row][offset_array_start + offset_equipment_name]
        quantity = schematic_tab[current_row][offset_array_start + offset_equipment_quantity]
        # if there is an equipment
        if equipment != "":
            # display it as many times as needed
            for ite in range(int(quantity)):
                # Get symbol height
                symbol_height = vs.HHeight(vs.GetObject(equipment))
                # Check if it is possible to create the object
                if room_info.labels_created:
                    # already existing room
                    # if there isn't enough space
                    if abs(page.y_pointer - symbol_height - 4) > page.drawing_area_height:
                        goto_next_page()
                else:
                    # new room
                    # 4 for end of room and 12 for room labels
                    if abs(page.y_pointer - symbol_height - 4 - 12) > page.drawing_area_height:
                        goto_next_page()
                # place symbol
                if not room_info.labels_created:
                    create_room_label(room_info.room, room_info.x + 4, room_info.y - 4)
                    create_room_label(room_info.floor, room_info.x + 44, room_info.y - 4)
                    page.y_pointer -= 12
                    room_info.height += 12
                    room_info.labels_created = 1
                vs.Symbol(equipment, (page.drawing_area_x + page.x, page.y + page.y_pointer), 0)
                if symbol_height % 4 == 0:  # To get an even number of spaces
                    page.y_pointer -= symbol_height + 4
                    room_info.height += symbol_height + 4
                else:
                    page.y_pointer -= symbol_height + symbol_height % 4
                    room_info.height += symbol_height + symbol_height % 4
                vs.SymbolToGroup(vs.LNewObj(), 0)
                vs.Ungroup()
        current_row += 1
    return current_row


def def_room_end(current_row, schematic_tab):
    global room_info, page
    vs.FillPat(0)
    vs.PenBack(colour_black)
    vs.PenFore(colour_black)
    vs.Rect((room_info.x, room_info.y), (room_info.x + room_info.width, room_info.y - room_info.height))
    page.y_pointer -= 4
    current_row += 1
    room_info.cable_number = 1
    vs.ForEachObject(replace_av_numbers, "S=AVTBC")
    vs.ForEachObject(replace_tv_numbers, "S=TVTBC")
    vs.ForEachObject(replace_local_numbers, "S=LocalTBC")
    vs.ForEachObject(replace_cctv_numbers, "S=CCTVTBC")
    vs.ForEachObject(replace_de_numbers, "S=DoorEntTBC")
    room_info.av_number += 1
    return current_row


def string_to_def(string):
    switcher = {
        "ROOM START": def_room_start,
        "ROOM EQUIPMENT START": def_equipment_start,
        "ROOM END": def_room_end
    }
    return switcher.get(string, "Invalid method")


# MAIN LOOP
fileName = vs.GetFile()
exSchematic = []
row = 0
with open(fileName) as csv_file:
    spam_reader = csv.reader(csv_file, delimiter=',')
    for row in spam_reader:
        exSchematic.append(row)
exSchematic.pop(0)

use_av_numbers = True
use_cable_numbers = True
use_tv_numbers = True
use_local_numbers = True
use_de_numbers = True
use_cctv_numbers = True
if use_cable_numbers and use_av_numbers:
    produce_excel_schedule = vs.YNDialog("Produce a cable schedule?")
vs.Symbol("Central System Rack Main", (page.x + 4, page.y - 4), 0)
for row in range(len(exSchematic)):
    # Get property string in first column of the schematic
    property_string = exSchematic[row][offset_array_start]
    # If no property string, then skip to next line
    if not property_string:
        continue
    # Otherwise, execute the associated method to the property string
    property_function = string_to_def(property_string)
    if property_function != "Invalid method":
        row = property_function(row, exSchematic)
if produce_excel_schedule:
    now = datetime.now()
    csvfile = open('Cable_Schedule_' + now.strftime("%d%m%Y_%H%M") + '.csv', 'w', newline='')
    # excel_stream = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL, dialect='excel')
    excel_stream = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, dialect='excel')
    excel_stream.writerow(['Floor', 'Room', 'Cable Type', 'AV Number', 'Cable Number', 'Purpose', 'From', 'To'])
    excel_stream.writerows(excel_rows)
    csvfile.close()
