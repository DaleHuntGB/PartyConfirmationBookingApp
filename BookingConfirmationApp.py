import tkinter as tk
from tkinter import ttk
import ttkthemes
import tkinter.simpledialog
import tkinter.filedialog
from tkcalendar import Calendar
import docx
from docx import Document
import json
from datetime import datetime
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
PARTY_COSTS = data["PARTY_COSTS"]

# Main Window
ApplicationWindow = tk.Tk()
ApplicationWindow.geometry(WINDOW_SIZE)
ApplicationWindow.title(f"{SITE_NAME} - Party Confirmation Booking")
# ApplicationWindow.iconbitmap("Media/FL_Logo.ico")
ApplicationWindow.resizable(False, False)

# Applying Theme
ApplicationTheme = ttkthemes.ThemedStyle(ApplicationWindow)
ApplicationTheme.set_theme("breeze")

# Global variables for entry widgets and other controls
nameInput, contactNumberInput, emailAddressInput = None, None, None
partyOptionsDropdown, partyFoodRoomDropdown, partyActivityRoomDropdown, partyDateSelector, partyStartTimeEntry, partyEndTimeEntry = None, None, None, None, None, None
staffNameInput, dateBookedSelector = None, None

# Create Menu
ApplicationMenu = tk.Menu(ApplicationWindow, tearoff=0)
ApplicationWindow.config(menu=ApplicationMenu)

# Add Rooms and Party Types Menu
SettingsMenu = tk.Menu(ApplicationMenu, tearoff=0)
ApplicationMenu.add_cascade(label="Settings", menu=SettingsMenu)
SettingsMenu.add_command(label="Site Name", command=lambda: UpdateSiteName())
SettingsMenu.add_command(label="Template Document", command=lambda: UpdateTemplateDocument())
SettingsMenu.add_separator()
SettingsMenu.add_command(label="Food Rooms", command=lambda: OpenFoodRoomsWindow())
SettingsMenu.add_command(label="Activity Rooms", command=lambda: OpenActivityRoomsWindow())
SettingsMenu.add_command(label="Party Types", command=lambda: OpenPartyTypesWindow())
SettingsMenu.add_command(label="Party Costs", command=lambda: OpenPartyCostsWindow())

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
    # ApplicationWindow.iconbitmap("Media/FL_Logo.ico")
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
    saveChangesButton = ttk.Button(ActivityRoomsWindow, text="Save Changes", command=lambda: SaveChanges(ActivityRoomsWindow))
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def OpenPartyTypesWindow():
    PartyTypesWindow = tk.Toplevel(ApplicationWindow)
    PartyTypesWindow.geometry("480x480")
    PartyTypesWindow.title("Update Party Types")
    # ApplicationWindow.iconbitmap("Media/FL_Logo.ico")
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
    saveChangesButton = ttk.Button(PartyTypesWindow, text="Save Changes", command=lambda: SaveChanges(PartyTypesWindow))
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def OpenFoodRoomsWindow():
    FoodRoomsWindow = tk.Toplevel(ApplicationWindow)
    FoodRoomsWindow.geometry("480x480")
    FoodRoomsWindow.title("Update Food Rooms")
    # ApplicationWindow.iconbitmap("Media/FL_Logo.ico")
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
    saveChangesButton = ttk.Button(FoodRoomsWindow, text="Save Changes", command=lambda: SaveChanges(FoodRoomsWindow))
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
        
        # Add Price
        priceLabel = ttk.Label(activityRoomSelectionWindow, text="Price", font=("Arial", 8, "bold underline"))
        priceLabel.pack(anchor="w", fill="x", padx=5, pady=3)
        priceInput = ttk.Entry(activityRoomSelectionWindow)
        priceInput.pack(anchor="w", fill="x", padx=5, pady=3)

        def saveActivityRooms():
            selectedActivityRooms = [room for room, var in activityRoomVars.items() if var.get()]
            PARTY_TYPES[partyTypeName] = selectedActivityRooms
            listbox.insert("end", partyTypeName)
            PARTY_COSTS[partyTypeName] = priceInput.get()
            UpdateDropdowns()
            activityRoomSelectionWindow.destroy()

        saveButton = ttk.Button(activityRoomSelectionWindow, text="Save", command=saveActivityRooms)
        saveButton.pack(anchor="center", pady=10)

def RemovePartyType(listbox):
    selectedPartyTypes = listbox.curselection()
    for partyIndex in selectedPartyTypes:
        partyTypeName = listbox.get(partyIndex)
        listbox.delete(partyIndex)
        del PARTY_COSTS[partyTypeName]
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

def SaveChanges(window):
    with open('BookingData.json', 'w') as file:
        json.dump(data, file, indent=4)
    UpdateDropdowns()
    window.destroy()

def UpdateDropdowns():
    partyOptionsDropdown["values"] = list(PARTY_TYPES.keys())
    partyActivityRoomDropdown["values"] = []

def OpenPartyCostsWindow():
    PartyCostsWindow = tk.Toplevel(ApplicationWindow)
    PartyCostsWindow.geometry("480x480")
    PartyCostsWindow.title("Update Party Costs")
    # ApplicationWindow.iconbitmap("Media/FL_Logo.ico")
    PartyCostsWindow.resizable(False, False)

    # Activity Rooms Heading
    headingLabel = ttk.Label(PartyCostsWindow, text="Party Costs", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")

    # Party Costs Listbox
    partyCostsListbox = tk.Listbox(PartyCostsWindow, selectmode="multiple")
    partyCostsListbox.pack(anchor="center", fill="both", expand=True, padx=5, pady=5)
    partyCostsListbox.config(selectmode="single")
    for party, cost in PARTY_COSTS.items():
        partyCostsListbox.insert("end", f"{party} - £{cost}")

    # Edit Party Costs Button
    editPartyCostsButton = ttk.Button(PartyCostsWindow, text="Edit Party Costs", command=lambda: EditPartyCosts(partyCostsListbox))
    editPartyCostsButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Delete Party Costs Button
    deletePartyCostsButton = ttk.Button(PartyCostsWindow, text="Delete Party Costs", command=lambda: RemovePartyCost(partyCostsListbox))
    deletePartyCostsButton.pack(anchor="center", fill="x", padx=5, pady=5)

    # Save Changes Button
    saveChangesButton = ttk.Button(PartyCostsWindow, text="Save Changes", command=lambda: SaveChanges(PartyCostsWindow))
    saveChangesButton.pack(anchor="center", fill="x", padx=5, pady=5)

def RemovePartyCost(listbox):
    selectedPartyCosts = listbox.curselection()
    for partyCostIndex in selectedPartyCosts:
        partyCost = listbox.get(partyCostIndex)
        partyName, cost = partyCost.split(" - £")
        del PARTY_COSTS[partyName]
        listbox.delete(partyCostIndex)

def EditPartyCosts(listbox):
    selectedPartyCosts = listbox.curselection()
    for partyCostIndex in selectedPartyCosts:
        partyCost = listbox.get(partyCostIndex)
        partyName, cost = partyCost.split(" - £")
        newCost = tk.simpledialog.askstring("Edit Party Cost", f"Enter New Cost for {partyName}")
        if newCost:
            PARTY_COSTS[partyName] = newCost
            listbox.delete(partyCostIndex)
            listbox.insert(partyCostIndex, f"{partyName} - £{newCost}")

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
    global staffNameInput, dateBookedSelector
    headingLabel = ttk.Label(ApplicationWindow, text="Admin Information", font=("Arial", 16, "bold"))
    headingLabel.pack(anchor="center")
    staffNameLabel = ttk.Label(ApplicationWindow, text="Staff Name", font=("Arial", 8, "bold underline"))
    staffNameLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    staffNameInput = ttk.Entry(ApplicationWindow)
    staffNameInput.pack(anchor="w", fill="x", padx=5, pady=3)
    dateBookedLabel = ttk.Label(ApplicationWindow, text="Date Booked", font=("Arial", 8, "bold underline"))
    dateBookedLabel.pack(anchor="w", fill="x", padx=5, pady=3)
    dateBookedSelector = Calendar(ApplicationWindow, selectmode="day")
    dateBookedSelector.pack(anchor="w", fill="x", padx=5, pady=3)

def GenerateDocument():
    global TEMPLATE_DOCUMENT
    shortenedCustomerName = nameInput.get().split(" ")[0]
    UKPartyDate = datetime.strptime(partyDateSelector.get_date(), "%m/%d/%y").strftime("%d/%m/%Y")
    UKBookedDate = datetime.strptime(dateBookedSelector.get_date(), "%m/%d/%y").strftime("%d/%m/%Y")
    
    CUSTOMER_INFORMATION = {
        "CUSTOMER_NAME": nameInput.get(),
        "FIRST_NAME": shortenedCustomerName,
        "CONTACT_NUMBER": contactNumberInput.get(),
        "EMAIL_ADDRESS": emailAddressInput.get()
    }
    
    PARTY_INFORMATION = {
        "PARTY_TYPE": partyOptionsDropdown.get(),
        "PARTY_FOOD_ROOM": partyFoodRoomDropdown.get(),
        "PARTY_ACTIVITY_ROOM": partyActivityRoomDropdown.get(),
        "PARTY_DATE": UKPartyDate,
        "PARTY_START_TIME": partyStartTimeEntry.get(),
        "PARTY_END_TIME": partyEndTimeEntry.get(),
        "COST_OF_PARTY": PARTY_COSTS[partyOptionsDropdown.get()]
    }

    ADMIN_INFORMATION = {
        "STAFF_NAME": staffNameInput.get(),
        "BOOKING_DATE": UKBookedDate
    }

    templateDocument = Document(TEMPLATE_DOCUMENT)

    for paragraph in templateDocument.paragraphs:
        for key, value in CUSTOMER_INFORMATION.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    # Check Tables
    for table in templateDocument.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in CUSTOMER_INFORMATION.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)
                for key, value in PARTY_INFORMATION.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)
                for key, value in ADMIN_INFORMATION.items():
                    if key in cell.text:
                        cell.text.replace(key, value)

    # Check Table Columns and Row Cells
    for table in templateDocument.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
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

    output_filename = f"Booking Confirmation - {CUSTOMER_INFORMATION['CUSTOMER_NAME']} - {PARTY_INFORMATION['PARTY_TYPE']}.docx"
    templateDocument.save(output_filename)
    tk.messagebox.showinfo("Success", f"Document saved as {output_filename}")

submitButton = ttk.Button(ApplicationWindow, text="Generate Party Confirmation", command=GenerateDocument)
submitButton.pack(anchor="center", side="bottom", padx=10, pady=5, fill="x")

GenerateCustomerInformationSection()
GeneratePartyInformationSection()
GenerateAdminSection()
ApplicationWindow.mainloop()
