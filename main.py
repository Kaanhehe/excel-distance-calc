import sys
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

### VARIABLES ###
try:
    starting_point = sys.argv[1]
    Distance_column = sys.argv[2]
    Adresse_column_Label = sys.argv[3]
except IndexError:
    starting_point = input("Enter starting point: ")
    Distance_column = ("Enter distance column: ")
    Adresse_column_Label = input("Enter address column label: ")

# Read the Excel file
df = pd.read_excel('Mappe1.xlsx')
wb = load_workbook('Mappe1.xlsx')
ws = wb.active

def calculate_distance(starting_point, destination_addresses):
    api_key = "AIzaSyDob6sH2CFjOYFn3mTO1FEB15nxmShZnx0"  # Replace with your own Google Maps API key
    url = "https://maps.googleapis.com/maps/api/distancematrix/json"

    params = {
        "origins": starting_point,
        "destinations": "|".join(destination_addresses),
        "key": api_key
    }

    response = requests.get(url, params=params)
    data = response.json()

    distances = []
    for i, row in enumerate(data["rows"]):
        for j, element in enumerate(row["elements"]):
            distance = element["distance"]["text"]
            distances.append(distance)
    return distances[0] if distances else None

# Calculate the distance for each row
df['Distance'] = df.apply(lambda row: pd.Series([calculate_distance(starting_point, [row[Adresse_column_Label]])]), axis=1)

# Insert the calculated distance into the Distance column
for cell in ws[Distance_column]: # type: ignore
    if cell.row > 1:
        cell.value = df.loc[cell.row - 2, 'Distance']

wb.save('Mappe1.xlsx')

# Print the DataFrame to the console
print("DataFrame after distance calculation:")
print(df)