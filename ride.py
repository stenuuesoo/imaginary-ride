import pandas as pd
import random
from datetime import datetime
import heapq


file_path = "keskused.xlsx"

df = pd.read_excel(file_path, engine='openpyxl')

data_dict = {}

for index, row in df.iterrows():
    data_dict[index] = row.to_dict()

first_point = 'Endla 47, 10615, Tallinn'
min_length = 1117
max_length = 1190

mileages = []  # set of 12 mileages that
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
    candidates = []

    for _ in range(10):  # Generate 10 candidates
        mileage_candidate = 0
        visited_locations_candidate = set()

        while mileage_candidate < min_length:
            row_number = random.choice(list(data_dict.keys()))
            if row_number not in visited_locations_candidate:
                center_data = data_dict[row_number]
                added_mileage = 2 * center_data['Kaugus aadressilt (km)']
                visited_locations_candidate.add(row_number)

                if (mileage_candidate + added_mileage) <= max_length:
                    mileage_candidate += added_mileage
                else:
                    break

        # Add the candidate mileage and visited_locations to the candidates list
        candidates.append((mileage_candidate, visited_locations_candidate))

    # Select the smallest mileage variant over min_length or the second best
    candidates.sort(key=lambda x: x[0])  # Sort candidates by mileage
    best_candidate = None
    for candidate in candidates:
        if candidate[0] >= min_length:
            best_candidate = candidate
            break

    if best_candidate is None:
        # If no candidate meets the condition, select the second best
        best_candidate = heapq.nsmallest(2, candidates, key=lambda x: x[0])[-1]

    mileage, visited_locations = best_candidate

    # Print the output for the best candidate
    for row_number in visited_locations:
        center_data = data_dict[row_number]
        print(
            f"{month_name}: {first_point} - {center_data['Keskuse nimi']} - {center_data['Aadress']} - {first_point} | {mileage:.1f} km | {start_number + mileage:.1f}")

    return mileage  # Return the mileage

# Call the do_work function
month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

for month_name in month_names:
    print("--------------------------------------------------")  # Add this line to print dashes between month changes
    monthly_mileage = do_work(month_name)
    mileages.append(monthly_mileage)

print("--------------------------------------------------")
print("Mileages:", ["{:.2f}".format(m) for m in mileages])
