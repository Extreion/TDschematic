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

# Page sizes
page_width = 288
page_height = 200


def create_room_label(value, x_l, y_l):
    vs.BeginGroup()
    vs.FillPat(1)
    vs.FillBack(colour_white)
    vs.PenBack(colour_white)
    vs.PenFore(colour_black)
    # Rounded rectangle
    vs.RRect((x_l, y_l), (x_l+rl_width, y_l-rl_height), rl_diamX, rl_diamY)
    # Line weight
    vs.SetLW(vs.LNewObj(), 10)
    # Create text
    vs.TextOrigin((x_l+rl_width/2, y_l-0.6))
    vs.TextFont(arial)
    vs.TextSize(text_size)
    vs.CreateText(value)
    vs.SetTextJust(vs.LNewObj(), 2)
    vs.EndGroup()


def def_room_start(current_row, schematic_tab):
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y
    global room_name, room_floor
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    current_row += 1
    room_y_pointer = room_y_pointer - 4
    start_of_room_x = current_page_x + 60
    start_of_room_y = room_y_pointer + start_of_room_y
    while schematic_tab[current_row][offset_array_start] != "ROOM STOP":
        property = schematic_tab[current_row][offset_array_start]
        value = schematic_tab[current_row][offset_array_start + offset_value]
        if property == "ROOM FLOOR":
            create_room_label(value, start_of_room_x + 4, start_of_room_y - 4)
            room_floor = value
        if property == "ROOM NAME":
            create_room_label(value, start_of_room_x + 44, start_of_room_y - 4)
            room_name = value
        current_row += 1
    return current_row


def def_equipment_start(current_row, schematic_tab):
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    current_row += 1
    room_y_pointer -= 8
    while schematic_tab[current_row][offset_array_start] != "ROOM EQUIPMENT END":
        equipment = schematic_tab[current_row][offset_array_start + offset_equipment_name]
        quantity = schematic_tab[current_row][offset_array_start + offset_equipment_quantity]
        if equipment != "":
            for ite in range(int(quantity)):
                vs.Symbol(equipment, (start_of_room_x, start_of_room_y + room_y_pointer), 0)
                #
                if vs.HHeight(vs.LNewObj()) % 4 == 0:
                    room_y_pointer -= vs.HHeight(vs.LNewObj()) + 4
                else:
                    room_y_pointer -= vs.HHeight(vs.LNewObj()) + vs.HHeight(vs.LNewObj()) % 4
                vs.SymbolToGroup(vs.LNewObj(), 0)
                vs.Ungroup()
        current_row += 1
    return current_row


def def_room_end(current_row, schematic_tab):
    global start_of_room_x, start_of_room_y, end_of_room_x, end_of_room_y
    global room_y_pointer, room_x_pointer  # Pointer to the current location of the equipment
    end_of_room_x = start_of_room_x + room_width
    end_of_room_y = start_of_room_y + room_y_pointer
    vs.FillPat(0)
    vs.PenBack(colour_black)
    vs.PenFore(colour_black)
    vs.Rect((start_of_room_x, start_of_room_y), (end_of_room_x, end_of_room_y))
    # Init variables for new room
    start_of_room_x = end_of_room_x - room_width
    start_of_room_y = end_of_room_y
    room_y_pointer = 0
    current_row += 1
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


for row in range(len(schematic)):
    # Get property string in first column of the schematic
    property_string = schematic[row][offset_array_start]
    # If no property string, then skip to next line
    if not property_string:
        continue
    # Otherwise, execute the associated method to the property string
    #vs.AlrtDialog(property_string)
    property_function = string_to_def(property_string)
    #vs.AlrtDialog(property_function)
    if property_function != "Invalid method":
        row = property_function(row, schematic)


