"""
TODO: File pathing
    - check for xls files in dir
    - check for info match
TODO: Output
    - label day and date
    - error handling
TODO: distribution
    - Package as executable
        - pyinstaller may work
"""
import pandas as pd
import os
import time
from pathlib import Path

client = input('Please specify a client: \n')
month = input('Please enter the month: \n')
file_path = input('Input file path: \n').strip('"')

timer_start = time.perf_counter()

file_list = Path(file_path).glob('*.xlsx')

path = Path.home().joinpath('Documents')

os.chdir(path)

with open(f'{client.title()}-{month.title()}.txt', 'w') as ticket_data:
    # Creates new text file with naming convention <Client>-<user specified date range>.txt
    for file in file_list:
        ticket_data.write(f'{file.stem}\n\n')  # Prints file name without path or extension
        with pd.ExcelFile(file) as xls:  # creates pandas dataframe for each sheet specified in file
            num_sheets = len(xls.sheet_names)
            if num_sheets < 7:
                days = [pd.read_excel(xls, parse_dates=['Start Time:', 'End Time:', 'Time Worked:'])]
            else:
                days = [pd.read_excel(xls, day, parse_dates=['Start Time:', 'End Time:', 'Time Worked:'])
                        for day in range(7)]

            for day in days:
                try:
                    if 'Customer' in day.columns.values:
                        day = day[[
                            'Start Time:',
                            'End Time:',
                            'Customer',
                            'Ticket Number/Action:',
                            'Time Worked:']].dropna()
                    else:
                        day = day[[
                            'Start Time:',
                            'End Time:',
                            'Company',
                            'Ticket Number/Action:',
                            'Time Worked:']].dropna()
                        day.rename(columns={
                            'Company': 'Customer'
                        }, inplace=True)

                    day.rename(columns={
                        'Start Time:': 'start',
                        'End Time:': 'end',
                        'Ticket Number/Action:': 'tickets',
                        'Time Worked:': 'worked'
                    }, inplace=True)

                    day.worked = day.end - day.start

                    if not day.empty:
                        client_data = day['Customer'].str.contains(client.title())
                        ticket_data.write(
                            f'{client.upper()} {day.worked.sum()} \n\n' + day.loc[client_data].to_string(
                                columns=['tickets', 'worked'], index=False, header=False
                            ) + '\n\n')

                except AttributeError:
                    pass

timer_finish = time.perf_counter()

print(f'Job\'s done. in {round(timer_finish - timer_start, 2)} second(s)')
