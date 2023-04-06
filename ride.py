import pandas as pd
import random
from datetime import datetime

file_path = "keskused.xlsx"

df = pd.read_excel(file_path, engine='openpyxl')

data_dict = {}

for index, row in df.iterrows():
    data_dict[index] = row.to_dict()

first_point = 'Endla 47, 10615, Tallinn'
min_length = 1117
max_length = 1190

mileage = 0
start_number = 314092

visited_locations = set()

def create_mileage():
    global mileage
    while True:
        row_number = random.choice(list(data_dict.keys()))
        if row_number not in visited_locations:
            center_data = data_dict[row_number]
            added_mileage = 2 * center_data['Kaugus aadressilt (km)']
            visited_locations.add(row_number)
            return added_mileage, center_data

def do_work(month_name):
    global mileage, visited_locations
    mileage = 0  # Reset the mileage variable for each iteration
    attempts = 0  # Add a counter for attempts
    max_attempts = 1000  # Set a maximum limit for attempts

    while mileage < min_length and attempts < max_attempts:
        added_mileage, center_data = create_mileage()
        if (mileage + added_mileage) <= max_length:
            mileage += added_mileage

            print(f"{month_name}: {first_point} - {center_data['Keskuse nimi']} - {center_data['Aadress']} - {first_point} | {mileage:.1f} km | {start_number + mileage:.1f}")

            if mileage >= max_length:
                save_data(visited_locations)
                visited_locations.clear()
                break  # Exit the loop when successful
        else:
            attempts += 1  # Increment the counter if the location is not suitable
            if attempts >= max_attempts:
                print(f"Unable to find suitable locations for month {month_name} after {max_attempts} attempts. Current mileage: {mileage:.1f} km")
                save_data(visited_locations)
                visited_locations.clear()
                break

def save_data(locations):
    global start_number
    filename = f"data_set_{start_number}.txt"
    with open(filename, "w") as f:
        f.write("\n".join(str(l) for l in locations))
    start_number += 1

def print_all_data():
    current_time = datetime.now().strftime("%H:%M:%S")
    output_filename = f"{current_time}.txt"

    with open(output_filename, "w") as output_file:
        for i in range(1, 13):
            output_file.write(f"Month {i}:\n")
            for j in range(1, start_number - 314091):
                filename = f"data_set_{j}.txt"
                with open(filename, "r") as f:
                    lines = f.readlines()
                    for line in lines:
                        parts = line.strip().split(" - ")
                        output_file.write(f"{parts[1]} - {parts[2]} - {parts[3]} - {parts[0]}\n")
            output_file.write("\n")
# Call the do_work function
month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

for month_name in month_names:
    print("--------------------------------------------------")  # Add this line to print dashes between month changes
    do_work(month_name)
# Print all the data sets
print_all_data()