import tkinter as tk
from tkinter import ttk
import ttkthemes
import tkinter.simpledialog
import tkinter.filedialog
from tkcalendar import Calendar
import docx
from docx import Document
import json
from datetime import date
import sys 
import os

WINDOW_SIZE = "480x940"
WINDOW_TITLE = "Party Confirmation Booking"

def LoadJSONData():
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    json_file_path = os.path.join(application_path, 'BookingData.json')
    with open(json_file_path, 'r') as file:
        return json.load(file)

data = LoadJSONData()
SITE_NAME = data["SITE_NAME"]
TEMPLATE_DOCUMENT = data["TEMPLATE_DOCUMENT"]
ACTIVITY_ROOMS = data["ACTIVITY_ROOMS"]
FOOD_ROOMS = data["FOOD_ROOMS"]
PARTY_TYPES = data["PARTY_TYPES"]

# Main Window
ApplicationWindow = tk.Tk()
ApplicationWindow.geometry(WINDOW_SIZE)
ApplicationWindow.title(f"{SITE_NAME} - Party Confirmation Booking")
ApplicationWindow.iconbitmap("Media/FL_Logo.ico")
ApplicationWindow.resizable(False, False)

# Applying Theme
ApplicationTheme = ttkthemes.ThemedStyle(ApplicationWindow)
ApplicationTheme.set_theme("breeze")

# Global variables for entry widgets and other controls
nameInput, contactNumberInput, emailAddressInput = None, None, None
partyOptionsDropdown, partyFoodRoomDropdown, partyActivityRoomDropdown, partyDateSelector, partyStartTimeEntry, partyEndTimeEntry = None, None, None, None, None, None
staffNameInput, receiptNumberInput, dateSentEntry = None, None, None

# Create Menu
ApplicationMenu = tk.Menu(ApplicationWindow, tearoff=0)
ApplicationWindow.config(menu=ApplicationMenu)

# Add Rooms and Party Types Menu
SettingsMenu = tk.Menu(ApplicationMenu, tearoff=0)
ApplicationMenu.add_cascade(label="Settings", menu=SettingsMenu)
SettingsMenu.add_command(label="Update Site Name", command=lambda: UpdateSiteName())
SettingsMenu.add_command(label="Update Template Document", command=lambda: UpdateTemplateDocument())
SettingsMenu.add_separator()
SettingsMenu.add_command(label="Update Food Rooms", command=lambda: OpenFoodRoomsWindow())
SettingsMenu.add_command(label="Update Activity Rooms", command=lambda: OpenActivityRoomsWindow())
SettingsMenu.add_command(label="Update Party Types", command=lambda: OpenPartyTypesWindow())

ApplicationMenu.add_separator()
ApplicationMenu.add_command(label="Exit", command=ApplicationWindow.quit)

def UpdateSiteName():
    global SITE_NAME, data
    SITE_NAME = tk.simpledialog.askstring("Update Site Name", "Enter Site Name")
    if SITE_NAME:
        data["SITE_NAME"] = SITE_NAME
        with open('BookingData.json', 'w') as file:
            json.dump(data, file, indent=4)
        # Reload the JSON data
        data = LoadJSONData()
        SITE_NAME = data["SITE_NAME"]
    ApplicationWindow.title(f"{SITE_NAME} - Party Confirmation Booking")

def UpdateTemplateDocument():
    global TEMPLATE_DOCUMENT, data
    TEMPLATE_DOCUMENT = tkinter.filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if TEMPLATE_DOCUMENT:
        data["TEMPLATE_DOCUMENT"] = TEMPLATE_DOCUMENT
        with open('BookingData.json', 'w') as file:
            json.dump(data, file, indent=4)
        # Reload the JSON data
        data = LoadJSONData()
        TEMPLATE_DOCUMENT = data["TEMPLATE_DOCUMENT"]



def OpenActivityRoomsWindow():
    ActivityRoomsWindow = tk.Toplevel(ApplicationWindow)
    ActivityRoomsWindow.geometry("480x480")
    ActivityRoomsWindow.title("Update Activity Rooms")
    ActivityRoomsWindow.iconbitmap("Media/FL_Logo.ico")
    ActivityRoomsWindow.resizable(False, False)

    # Activity Rooms Heading
    headingLabel = ttk.Label(ActivityRoomsWindow, text="Activity Rooms", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")

    # Activity Rooms Listbox
    activityRoomsListbox = tk.Listbox(ActivityRoomsWindow, selectmode="multiple")
    activityRoomsListbox.pack(anchor="center", fill="both", expand=True, padx=5, pady=5)
    activityRoomsListbox.config(selectmode="single")
    for room in ACTIVITY_ROOMS.keys():
        activityRoomsListbox.insert("end", room)

    # Add Room Button
    addRoomButton = ttk.Button(ActivityRoomsWindow, text="Add Room", command=lambda: AddRoom(activityRoomsListbox))
    addRoomButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Remove Room Button
    removeRoomButton = ttk.Button(ActivityRoomsWindow, text="Remove Room", command=lambda: RemoveRoom(activityRoomsListbox))
    removeRoomButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Save Changes Button
    saveChangesButton = ttk.Button(ActivityRoomsWindow, text="Save Changes", command=lambda: SaveChanges())
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def OpenPartyTypesWindow():
    PartyTypesWindow = tk.Toplevel(ApplicationWindow)
    PartyTypesWindow.geometry("480x480")
    PartyTypesWindow.title("Update Party Types")
    PartyTypesWindow.iconbitmap("Media/FL_Logo.ico")
    PartyTypesWindow.resizable(False, False)

    # Party Types Heading
    headingLabel = ttk.Label(PartyTypesWindow, text="Party Types", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")

    # Party Types Listbox
    partyTypesListbox = tk.Listbox(PartyTypesWindow, selectmode="multiple")
    partyTypesListbox.pack(anchor="center", fill="both", expand=True, padx=5, pady=5)
    partyTypesListbox.config(selectmode="single")
    for party in PARTY_TYPES.keys():
        partyTypesListbox.insert("end", party)

    # Add Party Type Button
    addPartyTypeButton = ttk.Button(PartyTypesWindow, text="Add Party Type", command=lambda: AddPartyType(partyTypesListbox))
    addPartyTypeButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Remove Party Type Button
    removePartyTypeButton = ttk.Button(PartyTypesWindow, text="Remove Party Type", command=lambda: RemovePartyType(partyTypesListbox))
    removePartyTypeButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Save Changes Button
    saveChangesButton = ttk.Button(PartyTypesWindow, text="Save Changes", command=lambda: SaveChanges())
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def OpenFoodRoomsWindow():
    FoodRoomsWindow = tk.Toplevel(ApplicationWindow)
    FoodRoomsWindow.geometry("480x480")
    FoodRoomsWindow.title("Update Food Rooms")
    FoodRoomsWindow.iconbitmap("Media/FL_Logo.ico")
    FoodRoomsWindow.resizable(False, False)

    # Food Rooms Heading
    headingLabel = ttk.Label(FoodRoomsWindow, text="Food Rooms", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")

    # Food Rooms Listbox
    foodRoomsListbox = tk.Listbox(FoodRoomsWindow, selectmode="multiple")
    foodRoomsListbox.pack(anchor="center", fill="both", expand=True, padx=5, pady=5)
    foodRoomsListbox.config(selectmode="single")
    for food_room in FOOD_ROOMS:
        foodRoomsListbox.insert("end", food_room)

    # Add Food Room Button
    addFoodRoomButton = ttk.Button(FoodRoomsWindow, text="Add Food Room", command=lambda: AddFoodRoom(foodRoomsListbox))
    addFoodRoomButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Remove Food Room Button
    removeFoodRoomButton = ttk.Button(FoodRoomsWindow, text="Remove Food Room", command=lambda: RemoveFoodRoom(foodRoomsListbox))
    removeFoodRoomButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Save Changes Button
    saveChangesButton = ttk.Button(FoodRoomsWindow, text="Save Changes", command=lambda: SaveChanges())
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def AddRoom(listbox):
    roomName = tk.simpledialog.askstring("Add Room", "Enter Room Name")
    if roomName:
        foodRoomVars = {}
        foodRoomSelectionWindow = tk.Toplevel(ApplicationWindow)
        foodRoomSelectionWindow.geometry("300x400")
        foodRoomSelectionWindow.title("Select Food Rooms")
        foodRoomSelectionWindow.resizable(False, False)

        # Food Rooms Checkboxes
        for foodRoom in FOOD_ROOMS:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(foodRoomSelectionWindow, text=foodRoom, variable=var)
            chk.pack(anchor='w')
            foodRoomVars[foodRoom] = var

        def saveFoodRooms():
            selectedFoodRooms = [room for room, var in foodRoomVars.items() if var.get()]
            ACTIVITY_ROOMS[roomName] = selectedFoodRooms
            listbox.insert("end", roomName)
            UpdateDropdowns()
            foodRoomSelectionWindow.destroy()

        saveButton = ttk.Button(foodRoomSelectionWindow, text="Save", command=saveFoodRooms)
        saveButton.pack(anchor="center", pady=10)

def RemoveRoom(listbox):
    selectedRooms = listbox.curselection()
    for roomIndex in selectedRooms:
        roomName = listbox.get(roomIndex)
        listbox.delete(roomIndex)
        del ACTIVITY_ROOMS[roomName]
        for partyType in PARTY_TYPES.values():
            if roomName in partyType:
                partyType.remove(roomName)
    UpdateDropdowns()

def AddPartyType(listbox):
    partyTypeName = tk.simpledialog.askstring("Add Party Type", "Enter Party Type Name")
    if partyTypeName:
        activityRoomVars = {}
        activityRoomSelectionWindow = tk.Toplevel(ApplicationWindow)
        activityRoomSelectionWindow.geometry("300x400")
        activityRoomSelectionWindow.title("Select Activity Rooms")
        activityRoomSelectionWindow.resizable(False, False)

        # Activity Rooms Checkboxes
        for activityRoom in ACTIVITY_ROOMS.keys():
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(activityRoomSelectionWindow, text=activityRoom, variable=var)
            chk.pack(anchor='w')
            activityRoomVars[activityRoom] = var

        def saveActivityRooms():
            selectedActivityRooms = [room for room, var in activityRoomVars.items() if var.get()]
            PARTY_TYPES[partyTypeName] = selectedActivityRooms
            listbox.insert("end", partyTypeName)
            UpdateDropdowns()
            activityRoomSelectionWindow.destroy()

        saveButton = ttk.Button(activityRoomSelectionWindow, text="Save", command=saveActivityRooms)
        saveButton.pack(anchor="center", pady=10)

def RemovePartyType(listbox):
    selectedPartyTypes = listbox.curselection()
    for partyIndex in selectedPartyTypes:
        partyTypeName = listbox.get(partyIndex)
        listbox.delete(partyIndex)
        del PARTY_TYPES[partyTypeName]
    UpdateDropdowns()

def AddFoodRoom(listbox):
    foodRoomName = tk.simpledialog.askstring("Add Food Room", "Enter Food Room Name")
    if foodRoomName:
        FOOD_ROOMS.append(foodRoomName)
        listbox.insert("end", foodRoomName)
        UpdateDropdowns()

def RemoveFoodRoom(listbox):
    selectedFoodRooms = listbox.curselection()
    for foodRoomIndex in selectedFoodRooms:
        foodRoomName = listbox.get(foodRoomIndex)
        listbox.delete(foodRoomIndex)
        FOOD_ROOMS.remove(foodRoomName)
        for room, foodRooms in ACTIVITY_ROOMS.items():
            if foodRoomName in foodRooms:
                foodRooms.remove(foodRoomName)
    UpdateDropdowns()

def SaveChanges():
    with open('BookingData.json', 'w') as file:
        json.dump(data, file, indent=4)
    UpdateDropdowns()

def UpdateDropdowns():
    partyOptionsDropdown["values"] = list(PARTY_TYPES.keys())
    partyActivityRoomDropdown["values"] = []

def GenerateCustomerInformationSection():
    global nameInput, contactNumberInput, emailAddressInput
    headingLabel = ttk.Label(ApplicationWindow, text="Customer Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    nameLabel = ttk.Label(ApplicationWindow, text="Customer Name", font=("Arial", 8, "bold underline"))
    nameLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    nameInput = ttk.Entry(ApplicationWindow)
    nameInput.pack(anchor="w", fill="x", padx=5, pady=3)
    contactNumberLabel = ttk.Label(ApplicationWindow, text="Contact Number", font=("Arial", 8, "bold underline"))
    contactNumberLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    contactNumberInput = ttk.Entry(ApplicationWindow)
    contactNumberInput.pack(anchor="w", fill="x", padx=5, pady=3)
    emailAddressLabel = ttk.Label(ApplicationWindow, text="Email Address", font=("Arial", 8, "bold underline"))
    emailAddressLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    emailAddressInput = ttk.Entry(ApplicationWindow)
    emailAddressInput.pack(anchor="w", fill="x", padx=5, pady=3)

def PartyRoomAvailability(event):
    global partyFoodRoomDropdown, partyActivityRoomDropdown
    partyType = partyOptionsDropdown.get()
    partyActivityRoomDropdown.set("Select...")
    partyFoodRoomDropdown.set("Select...")
    if partyType in PARTY_TYPES:
        activityRooms = PARTY_TYPES[partyType]
        partyActivityRoomDropdown["values"] = activityRooms

def UpdateFoodRoomAvailability(event):
    global partyFoodRoomDropdown
    partyActivityRoom = partyActivityRoomDropdown.get()
    partyFoodRoomDropdown.set("Select...")
    if partyActivityRoom in ACTIVITY_ROOMS:
        foodRooms = ACTIVITY_ROOMS[partyActivityRoom]
        partyFoodRoomDropdown["values"] = foodRooms

def GeneratePartyInformationSection():
    global partyOptionsDropdown, partyFoodRoomDropdown, partyActivityRoomDropdown, partyDateSelector, partyStartTimeEntry, partyEndTimeEntry
    headingLabel = ttk.Label(ApplicationWindow, text="Party Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    partyOptionsLabel = ttk.Label(ApplicationWindow, text="Party Type", font=("Arial", 8, "bold underline"))
    partyOptionsLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyOptionsDropdown = ttk.Combobox(ApplicationWindow, values=list(PARTY_TYPES.keys()))
    partyOptionsDropdown.pack(anchor="w", fill="x", padx=5, pady=3)
    partyOptionsDropdown.set("Select...")
    partyOptionsDropdown.bind("<<ComboboxSelected>>", lambda event: PartyRoomAvailability(event))
    partyActivityRoomLabel = ttk.Label(ApplicationWindow, text="Party Activity Room", font=("Arial", 8, "bold underline"))
    partyActivityRoomLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyActivityRoomDropdown = ttk.Combobox(ApplicationWindow, values=[])
    partyActivityRoomDropdown.pack(anchor="w", fill="x", padx=5, pady=3)
    partyActivityRoomDropdown.set("Select...")
    partyActivityRoomDropdown.bind("<<ComboboxSelected>>", lambda event: UpdateFoodRoomAvailability(event))
    partyFoodRoomLabel = ttk.Label(ApplicationWindow, text="Party Food Room", font=("Arial", 8, "bold underline"))
    partyFoodRoomLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyFoodRoomDropdown = ttk.Combobox(ApplicationWindow, values=[])
    partyFoodRoomDropdown.pack(anchor="w", fill="x", padx=5, pady=3)
    partyFoodRoomDropdown.set("Select...")
    partyDateLabel = ttk.Label(ApplicationWindow, text="Date", font=("Arial", 8, "bold underline"))
    partyDateLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyDateSelector = Calendar(ApplicationWindow, selectmode="day")
    partyDateSelector.pack(anchor="w", fill="x", padx=5, pady=3)
    partyTimeFrame = ttk.Frame(ApplicationWindow)
    partyTimeFrame.pack(anchor="center", fill="x", padx=5, pady=3)
    partyStartTimeLabel = ttk.Label(partyTimeFrame, text="Start", font=("Arial", 8, "bold underline"))
    partyStartTimeLabel.pack(anchor="w", padx=5, pady=3, side="left")
    partyStartTimeEntry = ttk.Entry(partyTimeFrame, width=15)
    partyStartTimeEntry.pack(anchor="w", padx=5, pady=3, side="left")
    partyEndTimeLabel = ttk.Label(partyTimeFrame, text="End", font=("Arial", 8, "bold underline"))
    partyEndTimeLabel.pack(anchor="e", padx=5, pady=3, side="right")
    partyEndTimeEntry = ttk.Entry(partyTimeFrame, width=15)
    partyEndTimeEntry.pack(anchor="e", padx=5, pady=3, side="right")

def GenerateAdminSection():
    global staffNameInput, receiptNumberInput, dateSentEntry
    headingLabel = ttk.Label(ApplicationWindow, text="Admin Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    staffNameLabel = ttk.Label(ApplicationWindow, text="Staff Name", font=("Arial", 8, "bold underline"))
    staffNameLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    staffNameInput = ttk.Entry(ApplicationWindow)
    staffNameInput.pack(anchor="w", fill="x", padx=5, pady=3)
    receiptNumberLabel = ttk.Label(ApplicationWindow, text="Receipt Number", font=("Arial", 8, "bold underline"))
    receiptNumberLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    receiptNumberInput = ttk.Entry(ApplicationWindow)
    receiptNumberInput.pack(anchor="w", fill="x", padx=5, pady=3)
    dateSentLabel = ttk.Label(ApplicationWindow, text="Date Sent", font=("Arial", 8, "bold underline"))
    dateSentLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    dateSentEntry = ttk.Entry(ApplicationWindow)
    dateSentEntry.pack(anchor="w", fill="x", padx=5, pady=3)
    dateSentEntry.insert(0, date.today().strftime("%d/%m/%Y"))
    dateSentEntry.config(state="readonly")

def GenerateDocument():
    global TEMPLATE_DOCUMENT
    
    CUSTOMER_INFORMATION = {
        "CUSTOMER_NAME": nameInput.get(),
        "CONTACT_NUMBER": contactNumberInput.get(),
        "EMAIL_ADDRESS": emailAddressInput.get()
    }
    
    PARTY_INFORMATION = {
        "PARTY_TYPE": partyOptionsDropdown.get(),
        "PARTY_FOOD_ROOM": partyFoodRoomDropdown.get(),
        "PARTY_ACTIVITY_ROOM": partyActivityRoomDropdown.get(),
        "PARTY_DATE": partyDateSelector.get_date(),
        "PARTY_START_TIME": partyStartTimeEntry.get(),
        "PARTY_END_TIME": partyEndTimeEntry.get(),
    }

    ADMIN_INFORMATION = {
        "STAFF_NAME": staffNameInput.get(),
        "RECEIPT_NUMBER": receiptNumberInput.get(),
        "DATE_SENT": dateSentEntry.get()
    }

    templateDocument = Document(TEMPLATE_DOCUMENT)

    for paragraph in templateDocument.paragraphs:
        for key, value in CUSTOMER_INFORMATION.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
        for key, value in PARTY_INFORMATION.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
        for key, value in ADMIN_INFORMATION.items():
            if key in paragraph.text:
                paragraph.text.replace(key, value)

    # Clear Entry Widgets
    nameInput.delete(0, "end")
    contactNumberInput.delete(0, "end")
    emailAddressInput.delete(0, "end")
    partyOptionsDropdown.set("Select...")
    partyFoodRoomDropdown.set("Select...")
    partyActivityRoomDropdown.set("Select...")
    partyStartTimeEntry.delete(0, "end")
    partyEndTimeEntry.delete(0, "end")
    receiptNumberInput.delete(0, "end")

    output_filename = f"Booking Confirmation - {CUSTOMER_INFORMATION['CUSTOMER_NAME']} - {PARTY_INFORMATION['PARTY_TYPE']}.docx"
    templateDocument.save(output_filename)
    tk.messagebox.showinfo("Success", f"Document saved as {output_filename}")

submitButton = ttk.Button(ApplicationWindow, text="Generate Party Confirmation", command=GenerateDocument)
submitButton.pack(anchor="center", side="bottom", padx=10, pady=5, fill="x")

GenerateCustomerInformationSection()
GeneratePartyInformationSection()
GenerateAdminSection()
ApplicationWindow.mainloop()
