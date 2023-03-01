"""
This program is used to automatically generate lot numbers and profile lables for Green Fox Plastics and is property of Green Fox Plastics
Created by Xavier Angus

At this point, it is a frankenstien of the old program, thrown into a gui shell. 
"""
import os
import tkinter as tk
from tkinter import ttk
from tkinter import *
import datetime
from tkcalendar import DateEntry
from pathlib import Path
import xlsxwriter
import lot_number_functions

##########################################################
#TKinterStuff
# Set up the root window
root = tk.Tk()
root.title("GFP Lot Number Generator")
root.geometry("360x450")

# Declare string variable for date
dateStr = tk.StringVar()

# Create a combo box with color options
combo_box1 = ttk.Combobox(root, values=["White", "Yellow", "Light_Grey", "Weathered_Wood", "Dark_Grey",
                                          "Lime_Green", "Aruba_Blue", "Turf_Green", "Cherry_Wood", "Cardinal_Red",
                                          "Patriot_Blue", "Tudor_Brown", "Black"])
combo_box1.current(0)
combo_box1.config(font=("Helvetica", 12))

# Create a combo box with profile options
combo_box2 = ttk.Combobox(root, values=["1/2 x 8", "3/4 x 1-3/4", "3/4 x 2-5/8", "3/4 x 3-1/2", "3/4 x 5-1/2",
                                          "1 x 5-1/2", "1-1/8 x 3-1/2", "1-1/2 x 1-1/2", "1-1/2 x 2-1/2", "1-1/2 x 3-1/2",
                                          "1-1/2 x 5-1/2", "1-1/2 x 9-1/2", "2-1/2 x 2-1/2", "3-1/2 x 3-1/2", "Bench Frame"])
combo_box2.current(0)
combo_box2.config(font=("Helvetica", 12))

# Create labels for the objects
label1 = tk.Label(root, text="Select Color")
label1.config(font=("Helvetica", 12))
label2 = tk.Label(root, text="Select Profile")
label2.config(font=("Helvetica", 12))
label3 = tk.Label(root, text="Select Line Number")
label3.config(font=("Helvetica", 12))
label4 = tk.Label(root, text="Select Pallet Number")
label4.config(font=("Helvetica", 12))
label5 = tk.Label(root, text="Select Number of Pallets")
label5.config(font=("Helvetica", 12))
label6 = tk.Label(root, text="Select Date")
label6.config(font=("Helvetica", 12))

# Create a date picker widget
date_picker = DateEntry(root, date_pattern="yyyy-MM-dd",textvariable=dateStr)
date_picker.set_date(datetime.date.today())
date_picker.config(font=("Helvetica", 12))

# Create spinboxes
spinbox1 = Spinbox(root, from_=1, to=20, width=5)
spinbox1.config(font=("Helvetica", 12))
spinbox2 = Spinbox(root, from_=1, to=20, width=5)
spinbox2.config(font=("Helvetica", 12))
spinbox3 = Spinbox(root, from_=1, to=20, width=5)
spinbox3.config(font=("Helvetica", 12))

#Create labels that will show program outputs
output_label = tk.Label(root)
output_label.config(font=("Helvetica", 12))
lot_label = tk.Label(root)
lot_label.config(font=("Helvetica", 12))
lot_number = tk.Label(root)
lot_number.config(font=("Helvetica", 12))

#checkbox function for half day
half_day = False
def checkbox_clicked():
    global check_var, half_day
    if check_var.get() == 1:
        half_day = True
    else:
        half_day = False

#Create Checkbox
check_var = tk.IntVar()
checkbox = tk.Checkbutton(root, text="Half Day?", variable=check_var, command=checkbox_clicked)

###########################################################################
# Functions are not located inside lot_number_functions.py for program inoperatibility

###########################################################################
#checkbox to forward time by half day, makes 1 label for first day instead of 2



# Create a button to display the selected options
def push_the_button():
    # Declare the global variables
    global colorInput, profileInput, selected_date, line_number, pallet_number, num_of_pallets, dateString

    # Save the selected options as variables 
    colorInput = combo_box1.get()
    profileInput = combo_box2.get()
    selected_date = date_picker.get_date()
    line_number = spinbox1.get()
    pallet_number = spinbox2.get()
    num_of_pallets = spinbox3.get()
    dateString = selected_date.strftime("%m%d%y")

    # Output the selected options
    input_properties = "Color : " + colorInput + "  || Profile: " + profileInput + "\n" + " Date: " + dateString +  "  || Line: " + line_number
    output_label["text"] = input_properties
    output_label.config(font=("Helvetica", 12))

    # Assign color and profile
    assignedColor = lot_number_functions.assignColor(colorInput)
    colorBin, colorString = lot_number_functions.getColor(assignedColor)
    assignedProfile = lot_number_functions.assignProfile(profileInput)
    boardBin, boardString = lot_number_functions.getBoardSize(assignedProfile)
    lineBin = lot_number_functions.getLineNumber(line_number)
    prodDateBin, day, month, year = lot_number_functions.getProdDate(dateString)
    palletNum, howMany = lot_number_functions.getPalletNum(pallet_number, num_of_pallets)

    #creates labels and increments by day
    i = 1
    first_run = True
    while i <= int(howMany):
        j = 0
        while j <= 1 and i <= int(howMany):
            if first_run == True & half_day == True:
                halfDayNum = 1
            else:
                halfDayNum = 0
            #get the date set by user
            makeDate, makeDateBin = lot_number_functions.setDate(day, month, year) 
            
            #create the lot number in binary
            rawLotStr = colorBin + boardBin + lineBin + makeDateBin
            
            #encode the lot binary using charset
            encodedLot = lot_number_functions.encode(rawLotStr)
            
            #append the pallet number to "-"
            palletSig = '-' + palletNum 
            
            #append lot number and pallet number
            lotNum = encodedLot + palletSig
            
            #print results to text file
            lot_number_functions.printToFile(colorString, boardString, lineBin, makeDateBin, palletNum, lotNum)
            
            #create excel label file
            lot_number_functions.printOut(colorString, palletNum, boardString, lotNum, makeDate)
            
            #each time loop is ran, next pallet number is created, then turned to string to be appended to next label
            palletNum = int(palletNum) + 1
            palletNum = str(palletNum)

            #j makes the loop only be allowed to run twice before leaving loop to increment to next day
            #half day count is added only for first run if desired by user
            j = j + 1 + halfDayNum

            #Turns off first run
            first_run = False
            
            #i adds towards the total amount of pallets made, this ends the loops entirelty.
            i = i + 1
        #increment the day by 1 day, skipping weekends and other set days
        day, month, year = lot_number_functions.dateInc(day, month, year)

    # Output the result
    lot_out = r"File located at \\lcc-fsqb-01.lcc.local\Shares\Green Fox\QC\Labels"
    lot_label["text"] = lot_out
    lot_number["text"] = "Lot # : " + lotNum


button = tk.Button(root, text="Generate Lot Number", command=push_the_button)
button.config(font=("Helvetica", 12))

# Place the widgets in the main window
label1.pack()
combo_box1.pack()
label2.pack()
combo_box2.pack()
label3.pack()
spinbox1.pack()
label4.pack()
spinbox2.pack()
label5.pack()
spinbox3.pack()
label6.pack()
date_picker.pack()
checkbox.pack()
button.pack()
output_label.pack()
lot_label.pack()
lot_number.pack()


# Start the main event loop
root.mainloop()