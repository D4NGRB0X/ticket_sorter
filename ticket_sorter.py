'''
TODO: File pathing
    - check for xls files in dir
    - check for info match
    - run script on all files
TODO: Output
    - total hours for day per client
    - label day and date
TODO: distribution
    - Package as executable
'''
import pandas as pd
from pathlib import Path
import time


client = input('Please specify a client: \n')
month = input('Please enter the month: \n')
file_path = input('Input file path: \n').strip('"')

timer_start = time.perf_counter()

file_list = Path(file_path).glob('*.xlsx')


with open(f'{client.title()}-{month.title()}.txt', 'w') as ticket_data:
    for file in file_list:
        with pd.ExcelFile(file) as xls:
            days = [pd.read_excel(xls, day, parse_dates=['Start Time:', 'End Time:', 'Time Worked:'])
                    for day in range(7)]

            for day in days:

                try:
                    day = day[[
                        'Start Time:',
                        'End Time:',
                        'Customer',
                        'Ticket Number/Action:',
                        'Time Worked:']].dropna()

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
                            f'{client.upper()} {day.worked.sum()} \n' + day.loc[client_data].to_string(
                                columns=['tickets'], index=False, header=False
                            ) + '\n\n')

                except AttributeError:
                    pass
timer_finish = time.perf_counter()

print(f'Job\'s done. in {round(timer_finish - timer_start,2)} second(s)')

