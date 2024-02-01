import sys
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

### VARIABLES ###
try:
    starting_point = sys.argv[1]
    adresse_column = sys.argv[2]
    distance_column = sys.argv[3]
    table_row = sys.argv[4]
    file_name = sys.argv[5]
    sheet_name = sys.argv[6]
except IndexError:
    starting_point = input("Enter starting point: ") # In den Baumgärten 12, 63225 Langen (Hessen)
    adresse_column = input("Enter address column (A, B, C...): ")
    distance_column = input("Enter distance column(A, B, C...): ")
    table_row = int(input("Enter table row (1, 2, 3...): ")) or 1
    file_name = input("Enter file name(without .xlsx): ") + ".xlsx" or "Mappe1.xlsx"
    sheet_name = input("Enter sheet name (or leave it empty): ") or "Tabelle1"

# Read the Excel file
df = pd.read_excel(file_name, sheet_name=sheet_name)
wb = load_workbook(file_name)
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
df['Distance'] = df.apply(lambda row: pd.Series([calculate_distance(starting_point, [row.iloc[table_row]])]), axis=1) # type: ignore

# Insert the calculated distance into the Distance column
for cell in ws[distance_column]: # type: ignore
    if cell.row > 1:
        cell.value = df.loc[cell.row - 2, 'Distance']

wb.save('Mappe1.xlsx')

# Print the DataFrame to the console
print("DataFrame after distance calculation:")
print(df)