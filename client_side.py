import openpyxl
from openpyxl import Workbook
from datetime import datetime, timedelta
import os
import subprocess

# Define the file name
file_name = "hotel_guest_data.xlsx"

# Notification function
def send_notification(title, message):
    # Use AppleScript for macOS notification
    script = f'display notification "{message}" with title "{title}"'
    subprocess.run(["osascript", "-e", script])

# Function to get the next available room number
def get_next_room_number(sheet, max_rooms=200):
    room_numbers = [cell.value for cell in sheet['F'] if isinstance(cell.value, int)]
    last_room = max(room_numbers) if room_numbers else 0
    if last_room < max_rooms:
        return last_room + 1
    else:
        print("All rooms are currently occupied.")
        return None

# Load workbook if it exists, otherwise create a new one
if os.path.exists(file_name):
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
else:
    workbook = Workbook()
    sheet = workbook.active
    # Write the headers if creating a new workbook
    sheet.append([
        "Name", "Address", "Phone Number", "ID Proof", 
        "Number of People", "Room Number(s)", "Check-in Date", "Check-out Date"
    ])

# Gather data from the user
name = input("Enter your name: ")
address = input("Enter your address: ")
phone_number = input("Enter your phone number: ")
id_proof = input("Enter your ID proof (e.g., Aadhar, PAN): ")
number_of_people = int(input("Enter the number of people: "))

# Calculate the number of rooms required
rooms_required = (number_of_people + 2) // 3  # Each room holds up to 3 people

# Proceed only if there are enough rooms available
room_numbers = []
for _ in range(rooms_required):
    room_number = get_next_room_number(sheet)
    if room_number is None:
        print("Not enough rooms available for your group.")
        break
    room_numbers.append(room_number)

if len(room_numbers) == rooms_required:
    # Set check-in date to today's date and get duration for stay
    check_in_date = datetime.today().date()
    days_of_stay = int(input("Enter the number of days you will stay: "))
    check_out_date = check_in_date + timedelta(days=days_of_stay)

    # Write data to the sheet
    sheet.append([
        name, address, phone_number, id_proof, 
        number_of_people, ", ".join(map(str, room_numbers)), check_in_date, check_out_date
    ])

    # Adjust column width for readability
    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Get the column name
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        sheet.column_dimensions[col_letter].width = max_length + 2  # Adjust width with padding

    # Save the workbook
    workbook.save(file_name)
    print(f"Data written successfully to '{file_name}'. Rooms assigned: {', '.join(map(str, room_numbers))}")

    # Send notification
    send_notification("Room Assigned", f"Rooms {', '.join(map(str, room_numbers))} successfully assigned to {name}.")
else:
    print("Not enough rooms available.")
