import pandas as pd
import random

file_path = "/Users/ninja/Documents/accounting-ride-log/keskused.xlsx"

df = pd.read_excel(file_path, engine='openpyxl')

data_dict = {}

for index, row in df.iterrows():
    data_dict[index] = row.to_dict()

first_point = 'Endla 47, 10615, Tallinn'
min_length = 1117
max_length = 1190

mileage = 0
start_number = 314092

def create_mileage():
    global mileage
    row_number = random.choice(list(data_dict.keys()))
    center_data = data_dict[row_number]
    added_mileage = 2 * center_data['Kaugus aadressilt (km)']
    
    return added_mileage, center_data

def do_work():
    global mileage
    while mileage < min_length:
        added_mileage, center_data = create_mileage()
        if (mileage + added_mileage) <= max_length:
            mileage += added_mileage
            
            print(f"{first_point} - {center_data['Keskuse nimi']} - {center_data['Aadress']} - {mileage:.1f} km - {first_point}")
            
            # save the results as a set.
            # once max_lenght is reached, create a new set and stop if 12 sets is reached
            # 

# Call the do_work function
do_work()