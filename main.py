import glob
import os
import pandas as pd
from datetime import datetime, date, timedelta
import re

# Given a partial file name, returns the most recent matching csv file if it was modified today
def findCSV(partialName):
    matching_files = glob.glob(partialName)

    # If matching files found, return the most recent one if it was modified today
    if len(matching_files) > 0:
        sorted_files = sorted(matching_files, key=os.path.getmtime, reverse=True)

        most_recent_modification = os.path.getmtime(sorted_files[0])

        today = date.today().strftime('%Y-%m-%d')

        if datetime.fromtimestamp(most_recent_modification).date() == date.fromisoformat(today):
            return sorted_files[0]
        
    return None

def offlineReporting():
    print("Starting Offline Report...")

    offlineCSV = "%_Communications_*.csv"
    path = findCSV(offlineCSV) # Use function to locate appropriate comms data file

    if path:
        df = pd.read_csv(path)

        # Formatting for day, date only
        df['day'] = pd.to_datetime(df['day']).dt.date

        # Get entries from previous day + desired columns
        previous_day = datetime.now().date() - timedelta(days=1)
        df_selected = df[df['day'] == previous_day]

        selected_columns = ['Serial', 'Name', 'Company']
        df_selected = df_selected[selected_columns]

        # Remove anything after the last dash in name (to remove plug type)
        df_selected['Name'] = df_selected['Name'].apply(lambda name: re.sub(r'(?i)(\s-|-\s).*$', '', name))

        # Aggregate by serial and sort appropriately
        aggregated_df = df_selected.groupby('Serial').first().reset_index()
        aggregated_df = aggregated_df.sort_values(['Company', 'Serial'])

        # Output report to file, with the date info as first line
        output_file = 'OfflineChargerReport.txt'
        previous_day_formatted = previous_day.strftime("%m/%d/%Y")

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
        
        print("Offline Report Done.")
        
    else:
        print("Recent offline charger data not found.")


def errorReporting():
    print("Starting Error Report...")

    errorsCSV = "Error_Count_*.csv"
    path = findCSV(errorsCSV) # Use function to locate appropriate comms data file

    if path:
        df = pd.read_csv(path)

        # Filter to only include DCFCs
        df = df[df['type'] == 'DCFC']

        # Filter to only include entries from yesterday
        previous_day = datetime.utcnow().date() - timedelta(days=1)
        previous_day_str = previous_day.strftime('%Y-%m-%d')

        df = df[df['utcdate'].str.startswith(previous_day_str)]

        # Create a new dataframe with the desired headers: serial, name, company, and codes
        df = df[['serial', 'name', 'company', 'code']].copy()

        # Group the dataframe by serial and aggregate the codes into one string
        df['codes'] = df.groupby('serial')['code'].transform(lambda x: ','.join(x))
        df['codes'] = df['codes'].apply(lambda x: ','.join(set(str(x).split(','))))

        # Drop code column
        df = df.drop('code', axis=1)

        # Remove anything after the last dash in name (to remove plug type)
        df['name'] = df['name'].apply(lambda name: re.sub(r'(?i)(\s-|-\s).*$', '', name))

        # Remove duplicate entries
        df = df.drop_duplicates(keep = 'first')

        # Reset the index and sort by serial
        df.reset_index(drop=True, inplace=True)
        df = df.sort_values(['serial'])

        # Output
        output_file = 'ChargerErrorReport.txt'
        previous_day_str = previous_day.strftime("%m/%d/%Y")
        df.to_csv(output_file, index=False, sep='|')

        # Replace the separator in the saved file with " - " for better readability (workaround to single char sep)
        with open(output_file, 'r') as file:
            content = file.read()
            content = content.replace('|', ' - ')

        with open(output_file, 'w') as file:
            file.write(content)

        # Open the .txt file with Notepad
        os.system(f'notepad.exe {output_file}')
        print ("Error Report Done")
    
    else:
        print("Recent charger error data not found.")

def main():
    print ("Working...")
    offlineReporting()
    errorReporting()

    print("All Done.")

main()