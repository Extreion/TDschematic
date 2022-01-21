# Vectorworks AV Script
# Written by Samuel Sueur
# Made for Tillman Domotics LLP

import vs
import csv

# Variable used to offset the starting column of the schematic in case of new columns in the excel formatting
offset_array_start = 0
offset_value = 2  # for room properties
offset_equipment_name = 1  # for equipment name
offset_equipment_quantity = 2  # for equipment quantity

x = -1008
y = 500
temp_y = 0

current_page_x = -1008
current_page_y = 500

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

# Room variables
start_of_room_x = 0
start_of_room_y = current_page_y
end_of_room_x = 0
end_of_room_y = 0
room_x_pointer = 0
room_y_pointer = 0
room_name = ""
room_floor = ""
room_width = 164
room_max_height = 192
room_av_number = 1
room_cable_number = 1
room_labels_created = 0  # Flag to know if labels have been created
room_floor_x = 0  # X pos of room floor label
room_floor_y = 0  # Y pos of room floor label
room_name_x = 0  # X pos of room name label
room_name_y = 0  # Y pos of room name label
new_room = 0  # Flag to know if this is a new room or an old room


# Global variables for rooms
global_tv_number = 1
global_local_number = 1
global_cctv_number = 1
global_de_number = 1

# Page sizes
page_width = 288
page_height = 200

# Page variables
page_number = 1

# User variables
use_av_numbers = 0
use_cable_numbers = 0
use_tv_numbers = 0

# Excel file
excel_rows = [[]]
excel_stream = 0


def replace_tv_numbers(handle):
    global use_tv_numbers
    global global_tv_number
    if use_tv_numbers:
        if global_tv_number >= 100:
            tv_name = "TV" + str(global_tv_number)
        else:
            if 100 > global_tv_number >= 10:
                tv_name = "TV0" + str(global_tv_number)
            else:
                tv_name = "TV00" + str(global_tv_number)
        vs.SetHDef(handle, vs.GetObject(tv_name))
        global_tv_number += 1
    return


def replace_local_numbers(handle):
    global use_local_numbers
    global global_local_number
    if use_local_numbers:
        if global_local_number >= 100:
            local_name = "Local" + str(global_local_number)
        else:
            if 100 > global_local_number >= 10:
                local_name = "Local0" + str(global_local_number)
            else:
                local_name = "Local00" + str(global_local_number)
        vs.SetHDef(handle, vs.GetObject(local_name))
        global_local_number += 1
    return


def replace_av_numbers(handle):
    global room_av_number, room_cable_number
    global use_cable_numbers, use_av_numbers
    global excel_rows
    if use_av_numbers:
        if room_av_number >= 100:
            av_name = "AV" + str(room_av_number)
        else:
            if 100 > room_av_number >= 10:
                av_name = "AV0" + str(room_av_number)
            else:
                av_name = "AV00" + str(room_av_number)
        if produce_excel_schedule:
            cable_type = vs.GetRField(handle, 'Cable Type', 'Cable Type')
            excel_rows.append([room_floor, room_name, cable_type, av_name, room_cable_number,
                                            "AV headend", "Room"])
        vs.SetHDef(handle, vs.GetObject(av_name))
    if use_cable_numbers:
        if room_cable_number >= 100:
            cable_name = str(room_cable_number)
        else:
            if 100 > room_cable_number >= 10:
                cable_name = "0" + str(room_cable_number)
            else:
                cable_name = "00" + str(room_cable_number)
        label_pos_tuple = vs.GetSymLoc(handle)
        label_pos_list = list(label_pos_tuple)
        label_pos_list[0] += 8
        label_pos_tuple = tuple(label_pos_list)
        vs.Symbol(cable_name, label_pos_tuple, 0)
        room_cable_number += 1
    return


def replace_cctv_numbers(handle):
    global global_cctv_number
    global use_cctv_numbers
    global excel_rows
    if use_cctv_numbers:
        if produce_excel_schedule:
            cable_type = vs.GetRField(handle, 'Cable Type', 'Cable Type')
            excel_rows.append([room_floor, room_name, cable_type, "CCTV", global_cctv_number,
                               "AV headend", "Room"])
        vs.SetHDef(handle, vs.GetObject("CCTV"))
    if use_cable_numbers:
        if global_cctv_number >= 100:
            cable_name = str(global_cctv_number)
        else:
            if 100 > global_cctv_number >= 10:
                cable_name = "0" + str(global_cctv_number)
            else:
                cable_name = "00" + str(global_cctv_number)
        label_pos_tuple = vs.GetSymLoc(handle)
        label_pos_list = list(label_pos_tuple)
        label_pos_list[0] += 8
        label_pos_tuple = tuple(label_pos_list)
        vs.Symbol(cable_name, label_pos_tuple, 0)
        global_cctv_number += 1
    return


def replace_de_numbers(handle):
    global global_de_number
    global use_de_numbers
    global excel_rows
    if use_de_numbers:
        if produce_excel_schedule:
            cable_type = vs.GetRField(handle, 'Cable Type', 'Cable Type')
            excel_rows.append([room_floor, room_name, cable_type, "DE", global_de_number,
                               "AV headend", "Room"])
        vs.SetHDef(handle, vs.GetObject("Door Ent"))
    if use_cable_numbers:
        if global_de_number >= 100:
            cable_name = str(global_de_number)
        else:
            if 100 > global_cctv_number >= 10:
                cable_name = "0" + str(global_de_number)
            else:
                cable_name = "00" + str(global_de_number)
        label_pos_tuple = vs.GetSymLoc(handle)
        label_pos_list = list(label_pos_tuple)
        label_pos_list[0] += 8
        label_pos_tuple = tuple(label_pos_list)
        vs.Symbol(cable_name, label_pos_tuple, 0)
        global_de_number += 1
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
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y
    global room_name, room_floor, room_labels_created, room_floor_x, room_floor_y, room_name_x, room_name_y
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    global new_room
    current_row += 1
    room_y_pointer = room_y_pointer - 4
    start_of_room_x = current_page_x + 60
    start_of_room_y -= 4
    while schematic_tab[current_row][offset_array_start] != "ROOM STOP":
        new_room = 1
        room_labels_created = 0
        property = schematic_tab[current_row][offset_array_start]
        value = schematic_tab[current_row][offset_array_start + offset_value]
        if property == "ROOM FLOOR":
            room_floor_x = start_of_room_x + 4
            room_floor_y = start_of_room_y - 4
            room_floor = value
        if property == "ROOM NAME":
            room_name_x = start_of_room_x + 44
            room_name_y = start_of_room_y - 4
            room_name = value
        current_row += 1
    return current_row


def start_new_room(room_labels):
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y, x, y, current_page_x
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    global room_floor, room_name
    global page_number
    global new_room
    global temp_y
    # Draw rectangle for old room
    end_of_room_x = start_of_room_x + room_width
    # end_of_room_y = start_of_room_y + room_y_pointer - 4
    end_of_room_y = temp_y - 4
    vs.FillPat(0)
    vs.PenBack(colour_black)
    vs.PenFore(colour_black)
    new_room = 0
    if room_labels:  # If labels have been created then draw a rectangle around the room
        vs.Rect((start_of_room_x, start_of_room_y), (end_of_room_x, end_of_room_y + 4))
    # Init positions for new room
    vs.Symbol("Central System Rack Main", (start_of_room_x - 42, y - int(((page_number-1)/7))*page_height - 8), 0)
    page_number += 1
    if (page_number-1) % 7 == 0:
        start_of_room_x = x + 60
        current_page_x = -1008
    else:
        start_of_room_x = end_of_room_x - room_width + page_width
        current_page_x += page_width
    start_of_room_y = y - int(((page_number-1)/7))*page_height - 4
    room_y_pointer = -12

    # Generate room labels again
    create_room_label(room_floor, start_of_room_x + 4, start_of_room_y - 4)
    create_room_label(room_name, start_of_room_x + 44, start_of_room_y - 4)
    return


def def_equipment_start(current_row, schematic_tab):
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y, y
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    global room_labels_created
    global new_room
    global temp_y
    temp_y = room_floor_y - 8
    current_row += 1
    room_y_pointer -= 8
    while schematic_tab[current_row][offset_array_start] != "ROOM EQUIPMENT END":
        equipment = schematic_tab[current_row][offset_array_start + offset_equipment_name]
        quantity = schematic_tab[current_row][offset_array_start + offset_equipment_quantity]
        if equipment != "":
            for ite in range(int(quantity)):
                symbol_height = vs.HHeight(vs.GetObject(equipment))
                if abs(room_y_pointer - symbol_height) > room_max_height:
                    start_new_room(room_labels_created)
                    room_labels_created = 1
                else:
                    if not room_labels_created:
                        create_room_label(room_floor, room_floor_x, room_floor_y)
                        create_room_label(room_name, room_name_x, room_name_y)
                        temp_y = room_floor_y - 8
                        room_labels_created = 1
                if not new_room:
                    vs.Symbol(equipment, (start_of_room_x, start_of_room_y + room_y_pointer), 0)
                else:
                    vs.Symbol(equipment, (start_of_room_x, temp_y), 0)
                if symbol_height % 4 == 0:  # To get an even number of spaces
                    room_y_pointer -= symbol_height + 4
                    temp_y -= symbol_height + 4
                else:
                    room_y_pointer -= symbol_height + symbol_height % 4
                    temp_y -= symbol_height + symbol_height % 4
                vs.SymbolToGroup(vs.LNewObj(), 0)
                vs.Ungroup()
        current_row += 1
    return current_row


def def_room_end(current_row, schematic_tab):
    global room_av_number, room_cable_number
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    global temp_y
    end_of_room_x = start_of_room_x + room_width
    if new_room:
        end_of_room_y = temp_y
    else:
        end_of_room_y = start_of_room_y + room_y_pointer
    vs.FillPat(0)
    vs.PenBack(colour_black)
    vs.PenFore(colour_black)
    vs.Rect((start_of_room_x, start_of_room_y), (end_of_room_x, end_of_room_y))
    # Init variables for new room
    start_of_room_x = end_of_room_x - room_width
    start_of_room_y = end_of_room_y
    current_row += 1
    room_cable_number = 1
    vs.ForEachObject(replace_av_numbers, "S=AVTBC")
    vs.ForEachObject(replace_tv_numbers, "S=TVTBC")
    vs.ForEachObject(replace_local_numbers, "S=LocalTBC")
    vs.ForEachObject(replace_cctv_numbers, "S=CCTVTBC")
    vs.ForEachObject(replace_de_numbers, "S=DoorEntTBC")
    room_av_number += 1
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
schematic = []
row = 0
with open(fileName) as csv_file:
    spam_reader = csv.reader(csv_file, delimiter=',')
    for row in spam_reader:
        schematic.append(row)
schematic.pop(0)

use_av_numbers = vs.YNDialog("Use AV numbers?")
use_cable_numbers = vs.YNDialog("Use cable numbers?")
use_tv_numbers = vs.YNDialog("Use TV numbers?")
use_local_numbers = vs.YNDialog("Use local numbers ?")
use_de_numbers = vs.YNDialog("Use Door Entry cables numbers?")
use_cctv_numbers = vs.YNDialog("Use CCTV numbers?")
if use_cable_numbers and use_av_numbers:
    produce_excel_schedule = vs.YNDialog("Produce a cable schedule?")
for row in range(len(schematic)):
    # Get property string in first column of the schematic
    property_string = schematic[row][offset_array_start]
    # If no property string, then skip to next line
    if not property_string:
        continue
    # Otherwise, execute the associated method to the property string
    property_function = string_to_def(property_string)
    if property_function != "Invalid method":
        row = property_function(row, schematic)
vs.Symbol("Central System Rack Main", (start_of_room_x - 42, y - int(((page_number-1)/7))*page_height - 8), 0)
if produce_excel_schedule:
    csvfile = open('test.csv', 'w', newline='')
    excel_stream = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL, dialect='excel')
    excel_stream.writerow(['Floor', 'Room', 'Cable Type', 'AV Number', 'Cable Number', 'From', 'To'])
    excel_stream.writerows(excel_rows)
    csvfile.close()
