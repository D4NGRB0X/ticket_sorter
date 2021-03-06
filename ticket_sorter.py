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
import re
import pandas as pd
import os
import time
from pathlib import Path
from itertools import chain

file_path = input('Input file path: \n').strip('"')
month = input('Please enter the month: \n')
client = input('Please specify a client: \n')


timer_start = time.perf_counter()

file_list = Path(file_path).glob('*.xlsx')

path = Path.home().joinpath('Documents')

os.chdir(path)


def data_handling(client):

    for day in days:
        try:
            if 'Company' in day.columns.values:  # test for series header customer
                day.rename(columns={
                    'Company': 'Customer'
                }, inplace=True)
            day = day[[
                'Start Time:',
                'End Time:',
                'Customer',
                'Ticket Number/Action:',
                'Time Worked:']].dropna()

            # rename series headers to pandas/python friendly
            day.rename(columns={
                'Start Time:': 'start',
                'End Time:': 'end',
                'Customer': 'customer',
                'Ticket Number/Action:': 'tickets',
                'Time Worked:': 'worked'
            }, inplace=True)  # does swap directly

            day.worked = day.end - day.start  # calculates time worked per ticket

            if not day.empty:  # ignores sheets with no data
                client_data = day.customer.str.contains(client, flags=re.IGNORECASE)  # setup for client data
                # flag=re.IGNORECASE checks for is catch all for user input
                if not day.loc[client_data].empty:
                    total_time = day.loc[client_data].worked.sum().to_pytimedelta()
                    ticket_data.write(  # writes column of all tickets worked for specific client
                        f'{client.upper()} {str(total_time):.7} \n\n' + day.loc[client_data].to_string(
                            columns=['tickets'], index=False, header=False
                        ) + '\n\n')

        except AttributeError:
            pass


for file in file_list:
    with open(f'{file.stem}-{client.title()}-{month.title()}.txt', 'a+') as ticket_data:
        # creates new text file with naming convention <Client>-<user specified date range>.txt
        ticket_data.write(f'{file.stem}\n\n')  # prints file name without path or extension
        with pd.ExcelFile(file) as xls:  # creates pandas dataframe for each sheet specified in file
            num_sheets = len(xls.sheet_names)
            if num_sheets < 7:  # data collected on single sheet
                days = [pd.read_excel(xls, parse_dates=['Start Time:', 'End Time:', 'Time Worked:'])]
            else:  # data collected on multiple sheets
                days = [pd.read_excel(xls, day, parse_dates=['Start Time:', 'End Time:', 'Time Worked:'])
                        for day in range(7)]

            if client == 'all' or client == '':
                client = 'all'
                clients = pd.read_excel(xls, 'Client_List', header=None, index_col=None)
                client_list = list(chain.from_iterable(clients.values.tolist()))
                for client in client_list:
                    data_handling(client)
                client = 'all'  # reset client var to empty

            else:
                data_handling(client)

timer_finish = time.perf_counter()

print(f'Job\'s done. in {round(timer_finish - timer_start, 2)} second(s)')
print(f'Your file is saved to {path}')


if __name__ == "__main__":
    pass
