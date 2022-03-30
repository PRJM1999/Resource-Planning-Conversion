from openpyxl import load_workbook
import pandas as pd
import numpy as np
from datetime import date, timedelta

# Insert Excel Workbook address
wb = "Resource_Summary.xlsm"

# Add Worksheet name here that contains data
worksheet_name = 'Deliverables Overview'

# Read's Excel File
df = pd.read_excel(wb, sheet_name = worksheet_name)

# Removes null values and drop unwanted dates - adjust this to remove unwanted columns
# This script will only focus on values that fall within 2022
df = df.where(pd.notnull(df), None)
df = df.drop(columns=['Oct-21', 'Nov-21', 'Dec-21', 'Jan-23', 'Feb-23', 'Mar-23', 'Apr-23', 'May-23', 'Jun-23', 'Jul-23'])

# Drops jobs that do not have a staff number or a job number
for index, row in df.iterrows():
    if (df["Staff Number"][index] == None) or (df["Job Number"][index] == None):
        df = df.drop([index])

# Return number for month value in text
month_dict = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}

# Function to return all the mondays in a year
def allmondays(year):
   d = date(year, 1, 1)                    # January 1st
   d += timedelta(days = 7 - d.weekday())  # First Sunday
   while d.year == year:
      yield d
      d += timedelta(days = 7)

# Create an array contain all Monday values
monday_list = []

# Add all Monday dates to array
for d in allmondays(2022):
   monday_list.append(d)

# Create dataframe with JobCode, staffid and all Mondays in the calender year
new_df = pd.DataFrame(columns=['JobCode'] + ['staffid'] + monday_list)

# Used as index to loop through columns
num = 0

# This is inefficient with 3 for loops - find better method
# Main function to add values to new dataframe
for index, row in df.iterrows():
    
    # values will act as new row for each entry
    values = []
    
    # Append to array job number and staff number
    values.append(int(pd.to_numeric(df["Job Number"][index])))
    values.append(int(pd.to_numeric(df["Staff Number"][index])))
    
    # Loop through months
    for i in range(len(df.columns) - 4):

        # Get current month in numerical format
        month = month_dict[str(df.columns[i + 4][:3])]

        # Loop through % workload table values - column 2 onwards
        for value in new_df.columns[2:]:
            
            # if month matches current month value
            if value.month == month:
                
                # Get the % hourly values
                hours = pd.to_numeric(df.iloc[num, i + 4])

                # Convert value to weekly hours rather than %
                values.append(hours * 37.5)

    # Add values to dataframe        
    new_df.loc[len(new_df)] = values
    
    # Increase number to loop through next set of hourly values
    num += 1

# Save new dataframe as a csv file, run api function from it
new_df.to_csv("API_Sheet.csv")