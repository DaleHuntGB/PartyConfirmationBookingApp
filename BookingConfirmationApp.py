import tkinter as tk
from tkinter import ttk
import ttkthemes
import tkinter.simpledialog
from tkcalendar import Calendar
import docx
from docx import Document
import json
from datetime import date

WINDOW_SIZE = "480x940"
WINDOW_TITLE = "Party Confirmation Booking"
TEMPLATE_DOCUMENT = "PartyBookingConfirmationTemplate.docx"
SITE_NAME = "Portslade Sports Centre"

def LoadJSONData():
    with open('BookingData.json', 'r') as file:
        return json.load(file)

data = LoadJSONData()
ACTIVITY_ROOMS = data["ACTIVITY_ROOMS"]
FOOD_ROOMS = data["FOOD_ROOMS"]
PARTY_TYPES = data["PARTY_TYPES"]
PARTY_ROOM_AVAILABILITY = data["PARTY_ROOM_AVAILABILITY"]

# Main Window
ApplicationWindow = tk.Tk()
ApplicationWindow.geometry(WINDOW_SIZE)
ApplicationWindow.title(f"{SITE_NAME} - Party Confirmation Booking")
ApplicationWindow.iconbitmap("FL_Logo.ico")
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

# Add Rooms Menu
RoomsMenu = tk.Menu(ApplicationMenu, tearoff=0)
ApplicationMenu.add_cascade(label="Settings", menu=RoomsMenu)
RoomsMenu.add_command(label="Update Activity Rooms", command=lambda: OpenActivityRoomsWindow())

def OpenActivityRoomsWindow():
    ActivityRoomsWindow = tk.Toplevel(ApplicationWindow)
    ActivityRoomsWindow.geometry("480x480")
    ActivityRoomsWindow.title("Update Activity Rooms")
    ActivityRoomsWindow.iconbitmap("FL_Logo.ico")
    ActivityRoomsWindow.resizable(False, False)

    # Activity Rooms Heading
    headingLabel = ttk.Label(ActivityRoomsWindow, text="Activity Rooms", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")

    # Activity Rooms Listbox
    activityRoomsListbox = tk.Listbox(ActivityRoomsWindow, selectmode="multiple")
    activityRoomsListbox.pack(anchor="center", fill="both", expand=True, padx=5, pady=5)
    activityRoomsListbox.config(selectmode="single")
    for room in ACTIVITY_ROOMS:
        activityRoomsListbox.insert("end", room)

    # Add Room Button
    addRoomButton = ttk.Button(ActivityRoomsWindow, text="Add Room", command=lambda: AddRoom(activityRoomsListbox))
    addRoomButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Remove Room Button
    removeRoomButton = ttk.Button(ActivityRoomsWindow, text="Remove Room", command=lambda: RemoveRoom(activityRoomsListbox))
    removeRoomButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Save Changes Button
    saveChangesButton = ttk.Button(ActivityRoomsWindow, text="Save Changes", command=lambda: SaveChanges(activityRoomsListbox))
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def AddRoom(listbox):
    roomName = tk.simpledialog.askstring("Add Room", "Enter Room Name")
    if roomName:
        activityVars = {}
        activitySelectionWindow = tk.Toplevel(ApplicationWindow)
        activitySelectionWindow.geometry("300x400")
        activitySelectionWindow.title("Select Activities")
        activitySelectionWindow.resizable(False, False)

        # Activity Checkboxes
        for activity in PARTY_TYPES:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(activitySelectionWindow, text=activity, variable=var)
            chk.pack(anchor='w')
            activityVars[activity] = var

        def selectFoodRooms():
            selectedActivities = [activity for activity, var in activityVars.items() if var.get()]
            activitySelectionWindow.destroy()

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
                for activity in selectedActivities:
                    PARTY_ROOM_AVAILABILITY.setdefault(activity, {})
                    PARTY_ROOM_AVAILABILITY[activity][roomName] = selectedFoodRooms
                listbox.insert("end", roomName)
                ACTIVITY_ROOMS.append(roomName)
                foodRoomSelectionWindow.destroy()

            saveButton = ttk.Button(foodRoomSelectionWindow, text="Save", command=saveFoodRooms)
            saveButton.pack(anchor="center", pady=10)

        selectActivitiesButton = ttk.Button(activitySelectionWindow, text="Next", command=selectFoodRooms)
        selectActivitiesButton.pack(anchor="center", pady=10)

def RemoveRoom(listbox):
    # Remove Selected Room from Listbox and JSON Data File
    selectedRooms = listbox.curselection()
    for roomIndex in selectedRooms:
        roomName = listbox.get(roomIndex)
        listbox.delete(roomIndex)
        ACTIVITY_ROOMS.remove(roomName)
        for activity in PARTY_ROOM_AVAILABILITY:
            if roomName in PARTY_ROOM_AVAILABILITY[activity]:
                del PARTY_ROOM_AVAILABILITY[activity][roomName]

def SaveChanges(listbox):
    # Save Changes to JSON Data File
    with open('BookingData.json', 'w') as file:
        json.dump(data, file, indent=4)

def GenerateCustomerInformationSection():
    global nameInput, contactNumberInput, emailAddressInput
    # Customer Heading
    headingLabel = ttk.Label(ApplicationWindow, text="Customer Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    # Customer Name
    nameLabel = ttk.Label(ApplicationWindow, text="Customer Name", font=("Arial", 8, "bold underline"))
    nameLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    nameInput = ttk.Entry(ApplicationWindow)
    nameInput.pack(anchor="w", fill="x", padx=5, pady=3)
    # Customer Contact Number
    contactNumberLabel = ttk.Label(ApplicationWindow, text="Contact Number", font=("Arial", 8, "bold underline"))
    contactNumberLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    contactNumberInput = ttk.Entry(ApplicationWindow)
    contactNumberInput.pack(anchor="w", fill="x", padx=5, pady=3)
    # Customer Email Address
    emailAddressLabel = ttk.Label(ApplicationWindow, text="Email Address", font=("Arial", 8, "bold underline"))
    emailAddressLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    emailAddressInput = ttk.Entry(ApplicationWindow)
    emailAddressInput.pack(anchor="w", fill="x", padx=5, pady=3)

def PartyRoomAvailability(event):
    global partyFoodRoomDropdown, partyActivityRoomDropdown
    partyType = partyOptionsDropdown.get()
    partyActivityRoomDropdown.set("Select...")
    partyFoodRoomDropdown.set("Select...")
    partyActivityRoomDropdown["values"] = list(ACTIVITY_ROOMS)
    partyFoodRoomDropdown["values"] = list(FOOD_ROOMS)
    if partyType in PARTY_ROOM_AVAILABILITY:
        activityRooms = list(PARTY_ROOM_AVAILABILITY[partyType].keys())
        partyActivityRoomDropdown["values"] = activityRooms
        partyActivityRoomDropdown.bind("<<ComboboxSelected>>", UpdateFoodRoomAvailability)

def UpdateFoodRoomAvailability(event):
    global partyFoodRoomDropdown
    partyType = partyOptionsDropdown.get()
    partyActivityRoom = partyActivityRoomDropdown.get()
    partyFoodRoomDropdown.set("Select...")
    if partyActivityRoom in PARTY_ROOM_AVAILABILITY[partyType]:
        foodRooms = PARTY_ROOM_AVAILABILITY[partyType][partyActivityRoom]
        partyFoodRoomDropdown["values"] = foodRooms

def GeneratePartyInformationSection():
    global partyOptionsDropdown, partyFoodRoomDropdown, partyActivityRoomDropdown, partyDateSelector, partyStartTimeEntry, partyEndTimeEntry
    # Party Information
    headingLabel = ttk.Label(ApplicationWindow, text="Party Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    # Party Options
    partyOptionsLabel = ttk.Label(ApplicationWindow, text="Party Type", font=("Arial", 8, "bold underline"))
    partyOptionsLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyOptionsDropdown = ttk.Combobox(ApplicationWindow, values=PARTY_TYPES)
    partyOptionsDropdown.pack(anchor="w", fill="x", padx=5, pady=3)
    partyOptionsDropdown.set("Select...")
    partyOptionsDropdown.bind("<<ComboboxSelected>>", lambda event: PartyRoomAvailability(event))
    # Party Activity Room
    partyActivityRoomLabel = ttk.Label(ApplicationWindow, text="Party Activity Room", font=("Arial", 8, "bold underline"))
    partyActivityRoomLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyActivityRoomDropdown = ttk.Combobox(ApplicationWindow, values=ACTIVITY_ROOMS)
    partyActivityRoomDropdown.pack(anchor="w", fill="x", padx=5, pady=3)
    partyActivityRoomDropdown.set("Select...")
    # Party Food Room
    partyFoodRoomLabel = ttk.Label(ApplicationWindow, text="Party Food Room", font=("Arial", 8, "bold underline"))
    partyFoodRoomLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyFoodRoomDropdown = ttk.Combobox(ApplicationWindow, values=FOOD_ROOMS)
    partyFoodRoomDropdown.pack(anchor="w", fill="x", padx=5, pady=3)
    partyFoodRoomDropdown.set("Select...")
    # Party Date
    partyDateLabel = ttk.Label(ApplicationWindow, text="Date", font=("Arial", 8, "bold underline"))
    partyDateLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    partyDateSelector = Calendar(ApplicationWindow, selectmode="day")
    partyDateSelector.pack(anchor="w", fill="x", padx=5, pady=3)
    # Party Time Frame
    partyTimeFrame = ttk.Frame(ApplicationWindow)
    partyTimeFrame.pack(anchor="center", fill="x", padx=5, pady=3)
    # Party Time
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
    # Admin Heading
    headingLabel = ttk.Label(ApplicationWindow, text="Admin Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    # Staff Name
    staffNameLabel = ttk.Label(ApplicationWindow, text="Staff Name", font=("Arial", 8, "bold underline"))
    staffNameLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    staffNameInput = ttk.Entry(ApplicationWindow)
    staffNameInput.pack(anchor="w", fill="x", padx=5, pady=3)
    # Receipt Number
    receiptNumberLabel = ttk.Label(ApplicationWindow, text="Receipt Number", font=("Arial", 8, "bold underline"))
    receiptNumberLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    receiptNumberInput = ttk.Entry(ApplicationWindow)
    receiptNumberInput.pack(anchor="w", fill="x", padx=5, pady=3)
    # Date Sent
    dateSentLabel = ttk.Label(ApplicationWindow, text="Date Sent", font=("Arial", 8, "bold underline"))
    dateSentLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    dateSentEntry = ttk.Entry(ApplicationWindow)
    dateSentEntry.pack(anchor="w", fill="x", padx=5, pady=3)
    dateSentEntry.insert(0, date.today().strftime("%d/%m/%Y"))
    dateSentEntry.config(state="readonly")

def GenerateDocument():
    
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
                paragraph.text = paragraph.text.replace(key, value)

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

    templateDocument.save(f"Booking Confirmation - {nameInput.get()} - {partyOptionsDropdown.get()}.docx")

# Submit Button using ttk Button
submitButton = ttk.Button(ApplicationWindow, text="Generate Party Confirmation", command=GenerateDocument)
submitButton.pack(anchor="center", side="bottom", padx=10, pady=5, fill="x")

# Run Application
GenerateCustomerInformationSection()
GeneratePartyInformationSection()
GenerateAdminSection()
ApplicationWindow.mainloop()
