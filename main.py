import glob
import os
import pandas as pd
from datetime import datetime, date, timedelta
import re
import shutil

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

# Moves the error + comms logs from downloads folder to project folder
def moveFiles():
    df = pd.read_csv('paths.txt') # Get source + project paths
    source = df.iloc[0,0]
    destination = df.iloc[1,0]

    # Find comms log file
    commsCSV = "%_Communications_*.csv"
    commPath = os.path.join(source,commsCSV)
    commsFile = findCSV(commPath)

    # Find error log file
    errorCSV = "Error_Count_*.csv"
    errorPath = os.path.join(source,errorCSV)
    errorFile = findCSV(errorPath)

    if (commsFile):
        shutil.copy2(commsFile, destination)

    else:
        print("No recent comms log found in downloads.")

    if (errorFile):
        shutil.copy2(errorFile, destination)

    else:
        print("No recent error log found in downloads.")




def offlineReporting():
    print("Starting Offline Report...")

    excluded_serials = ['12000312', '01.O74-00157', 'KINT00016', 'ZEFDUMMY', 'ZEFDUMMY001'] # Serials to ignore (Mostly test units)

    offlineCSV = "%_Communications_*.csv"
    path = findCSV(offlineCSV) # Use function to locate appropriate comms data file

    if path:
        df = pd.read_csv(path)

        # Formatting for day, date only
        df['day'] = pd.to_datetime(df['day']).dt.date

        # Get entries from previous day + desired columns
        previous_day = datetime.now().date() - timedelta(days=1)
        df = df[df['day'] == previous_day]

        selected_columns = ['Serial', 'Name', 'Company']
        df = df[selected_columns]

        # Remove anything after the last dash in name (to remove plug type)
        df['Name'] = df['Name'].apply(lambda name: re.sub(r'(?i)(\s-|-\s).*$', '', name))

        # Aggregate by serial and sort appropriately
        df = df.groupby('Serial').first().reset_index()
        df = df.sort_values(['Company', 'Serial'])

        # Filter out excluded serials
        df = df[~df['Serial'].isin(excluded_serials)]

        # Output report to file, with the date info as first line
        output_file = 'OfflineChargerReport.txt'
        previous_day_formatted = previous_day.strftime("%m/%d/%Y")

        with open(output_file, 'w') as file:
            file.write(previous_day_formatted +" Offline Chargers:" + '\n')

        df.to_csv(output_file, index=False,header=False,mode='a', sep='|')

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

        # Filter to only include DCFCs with entries from yesterday
        df = df[df['type'] == 'DCFC']

        previous_day = datetime.utcnow().date() - timedelta(days=1)
        previous_day_str = previous_day.strftime('%Y-%m-%d')

        df = df[df['utcdate'].str.startswith(previous_day_str)]

        # Recreate with desired headers: serial, name, company, and code
        df = df[['serial', 'name', 'company', 'code']].copy()

        # Group the dataframe by serial and aggregate the codes into one column, drop individual code column
        df['codes'] = df.groupby('serial')['code'].transform(lambda x: ','.join(x))
        df['codes'] = df['codes'].apply(lambda x: ','.join(set(str(x).split(','))))

        df = df.drop('code', axis=1)

        # Remove anything after the last dash in name (to remove plug type)
        df['name'] = df['name'].apply(lambda name: re.sub(r'(?i)(\s-|-\s).*$', '', name))

        # Remove duplicate entries, reset index, sort by serial
        df = df.drop_duplicates(keep = 'first')
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
    moveFiles()
    
    offlineReporting()
    errorReporting()

    print("All Done.")

main()