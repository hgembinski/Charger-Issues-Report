import glob
import os
import pandas as pd
from datetime import datetime, date, timedelta
import re

# Given a partial file name, returns the most recent matching csv file if it was modified today
def findCSV(partialName):
    # Search for matching file names
    matching_files = glob.glob(partialName)

    # Check if any matching files were found
    if len(matching_files) > 0:
        # Sort the matching files by modification time in descending order
        sorted_files = sorted(matching_files, key=os.path.getmtime, reverse=True)
        # Get the modification time of the most recent file
        most_recent_modification = os.path.getmtime(sorted_files[0])
        # Get today's date
        today = date.today().strftime('%Y-%m-%d')
        # Check if the most recent file was modified today
        if datetime.fromtimestamp(most_recent_modification).date() == date.fromisoformat(today):
            # Return the most recent matching file
            return sorted_files[0]
        
    return None

def offlineReporting():
    print("Starting Offline Report...")
    offlineCSV = "%_Communications_*.csv"
    path = findCSV(offlineCSV) # Use function to locate appropriate comms data file

    if path:
        df = pd.read_csv(path)

        # Get the date of the previous day
        df['day'] = pd.to_datetime(df['day']).dt.date
        previous_day = datetime.now().date() - timedelta(days=1)

        # Get entries from previous day + desired columns
        df_selected = df[df['day'] == previous_day]

        selected_columns = ['Serial', 'Name', 'Company']
        df_selected = df_selected[selected_columns]

        # Remove anything after the last dash in name (to remove plug type)
        df_selected['Name'] = df_selected['Name'].apply(lambda name: re.sub(r'(?i)(\s-|-\s).*$', '', name))

        # Aggregate by serial and sort appropriately
        aggregated_df = df_selected.groupby('Serial').first().reset_index()
        aggregated_df = aggregated_df.sort_values(['Company', 'Serial'])
        
        # Format previous_day info to use as file header
        previous_day_formatted = previous_day.strftime("%m/%d/%Y")

        # Output report to file, with the date info as first line
        output_file = 'OfflineChargerReport.txt'

        with open(output_file, 'w') as file:
            file.write(previous_day_formatted +" Offline Chargers:" + '\n')

        aggregated_df.to_csv(output_file, index=False,header=False,mode='a', sep='|')

        # Replace the separator in the saved file with " - " for better readability (workaround to single char sep)
        with open(output_file, 'r') as file:
            content = file.read()
            content = content.replace('|', ' - ')

        with open(output_file, 'w') as file:
            file.write(content)

        # Open the .txt file with Notepad
        os.system(f'notepad.exe {output_file}')
        
    else:
        print("Recent offline charger data not found.")

    print("Offline Report Done.")

def errorReporting():

    print ("Error Report Done")

def main():
    print ("Working...")
    offlineReporting()
    errorReporting()

    print("All Done.")

main()