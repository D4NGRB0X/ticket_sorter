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
import os
from pathlib import Path


file_path = input('Input file path: \n')
month = input('Please enter the month: \n')

#location = os.chdir(file_path)

with pd.ExcelFile(file_path) as xls:
    sun = pd.read_excel(xls, 0)
    mon = pd.read_excel(xls, 1)
    tue = pd.read_excel(xls, 2)
    wed = pd.read_excel(xls, 3)
    thu = pd.read_excel(xls, 4)
    fri = pd.read_excel(xls, 5)
    sat = pd.read_excel(xls, 6)

days = [sun, mon, tue, wed, thu, fri, sat]
with open(f'{month.title()}.txt', 'w') as ticket_data:
    for day in days:
        try:
            day = day[['Customer', 'Ticket Number/Action:', 'Time Worked:']].dropna()
            if not day.empty:
                cl1 = day['Customer'].str.contains('Client1')
                ticket_data.write('Client 1 \n' + day.loc[cl1].to_string(
                    columns=['Ticket Number/Action:'], index=False, header=False
                ) + '\n\n')
                cl2 = day['Customer'].str.contains('Client2')
                ticket_data.write(
                    'Client 2 \n' + day.loc[cl2].to_string(
                        columns=['Ticket Number/Action:'], index=False, header=False
                    ) + '\n\n')

        except AttributeError:
            pass
print('Job\'s done.')
