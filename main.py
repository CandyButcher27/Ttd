import pandas as pd

# Set Pandas to display all rows and columns
pd.set_option('display.max_rows', None)  # Show all rows
pd.set_option('display.max_columns', None)  # Show all columns

#Reading the data from staff duties and staff leave sheets

file_path_1 = "staff duties.xlsx"
file_path_2 = "staff leave.xlsx"

room_data = pd.read_excel(file_path_1, sheet_name="ROOM")
staff_data = pd.read_excel(file_path_1, sheet_name="STAFF",header=None)
leave_data = pd.read_excel(file_path_2)

#Cleaning ROOM data

room_data[['Date','Start Time','End Time']] = room_data['Time'].str.split('|', expand = True)
room_data = room_data.drop(columns=['Time'])
room_data['Period'] = room_data['Start Time'].apply(lambda x : 'AN' if x == '09:30' else 'FN' if x=='14:00' else '')

def get_floor(room_name):
    if room_name[-3:].isdigit():
        floor_number = int(room_name[-3])
        return "Ground Floor" if floor_number==1 else "First Floor"
    return "Reserved"

room_data['Floor'] = room_data['Room'].apply(get_floor)


#Cleaning STAFF data
staff_data.columns=['SNO','ID','Name','Branch','Role','Mobile Number','Email']

#Merging Data (staff data and leave data)
merged_data = pd.merge(staff_data, leave_data[['Name', 'ID' , 'end_date']], on=['ID','Name'], how = 'left')

# Assinging Group captains and Room Captains
room_captains = merged_data[merged_data['Role']=="ROOM CAPTAIN"]
group_captains = merged_data[merged_data['Role']=="GROUP CAPTAIN"]

#Indexing Room and Group Captains according to their branch
group_captains = group_captains.set_index("Branch")
room_captains = room_captains.set_index("Branch") 

group_captains = group_captains.sort_index()
room_captains = room_captains.sort_index()  

room_data = room_data.sort_values(by=["Room","Date","Period"])
room_data = room_data.drop_duplicates()

#Allotment Logic

room_data['Date'] = pd.to_datetime(room_data['Date'], format='%d-%m-%y')
room_captains['end_date'] = pd.to_datetime(room_captains['end_date'], format='%d-%m-%y', errors='coerce')
group_captains['end_date'] = pd.to_datetime(group_captains['end_date'], format='%d-%m-%y', errors='coerce')

#Allotment of room captains
def allot_room_captains(room_data, room_captains):
    room_data['Room Captain'] = None
    duties = {captain: [] for captain in room_captains['ID']}

    for idx, row in room_data.iterrows():
        available_captains = room_captains[room_captains['end_date'].isna() | (room_captains['end_date'] != row['Date'])]
        assigned_captains = []

        for _, captain_row in available_captains.iterrows():
            captain_id = captain_row['ID']
            captain_name = captain_row['Name']
            if (len(duties[captain_id]) < 10 and
                not any(duty_date == row['Date'] and duty_period != row['Period'] for duty_date, duty_period in duties[captain_id])):

                assigned_captains.append(f"{captain_id} - {captain_name}")
                duties[captain_id].append((row['Date'], row['Period']))

                if row['Room'] in ['F102', 'F105'] and len(assigned_captains) < 2:
                    continue
                else:
                    break

        room_data.at[idx, 'Room Captain'] = ', '.join(assigned_captains)

    # Convert date back to desired format for display
    room_data['Date'] = room_data['Date'].dt.strftime('%d-%m-%Y')

    return room_data

room_data = allot_room_captains(room_data, room_captains)


