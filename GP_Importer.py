import os
import sys
import tkinter as tk
from tkinter.filedialog import askopenfilename
import csv
import openpyxl
import pandas as pd
from datetime import *


def main():
    root = tk.Tk(className=' GP Converter')
    w = 350
    h = 235
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)

    root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    # Add user input direction
    instruction = tk.Label(root, text='Include Project Size?\nNote: If Before 2pm CST then select no.')
    instruction.pack(pady=(30, 0))

    include_size = tk.BooleanVar()
    radiobutton_1 = tk.Radiobutton(root, text='Yes', variable=include_size, value=True)
    radiobutton_1.pack(pady=(15, 0))
    radiobutton_2 = tk.Radiobutton(root, text='No', variable=include_size, value=False)
    radiobutton_2.pack()

    def submit_function():
        print('Include Size:', include_size.get())
        root.destroy()

    submit = tk.Button(root, text='Next', command=submit_function)
    submit.pack(pady=10)

    exit_button = tk.Button(root, text='Quit', command=sys.exit)
    exit_button.pack()

    root.mainloop()
    tk.Tk().withdraw()  # We don't want a full GUI, so keep the root window from appearing
    input_file = askopenfilename(title='Select \'GP Project Data\' file')

    # def task():
    #     try:
    #         gp_convert(input_file, include_size.get())
    #     except MissingColumnException as ex:
    #         root.destroy()
    #         root1 = tk.Tk(className=' Task Failed')
    #         root1.geometry('%dx%d+%d+%d' % (w, h, x, y))
    #         error_message = tk.Label(root1, text=str(ex))
    #         error_message.pack(pady=(30, 0))
    #         exit_button = tk.Button(root1, text='Quit', command=sys.exit)
    #         exit_button.pack(pady=(15, 0))
    #         root1.mainloop()
    #     except Exception as ex:
    #         root.destroy()
    #         root1 = tk.Tk(className=' Task Failed')
    #         root1.geometry('%dx%d+%d+%d' % (w, h, x, y))
    #         error_message = tk.Label(root1, text='Unexpected Error: ' + str(ex))
    #         error_message.pack(pady=(30, 0))
    #         exit_button = tk.Button(root1, text='Quit', command=sys.exit)
    #         exit_button.pack(pady=(15, 0))
    #         root1.mainloop()
    #     root.destory()
    #
    # root = tk.Tk(className=' Task in Progress')
    # root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    # message = tk.Label(root, text='This may take up to 60 seconds.')
    # message.pack(pady=(95, 0))
    # root.after(200, task)
    # root.mainloop()

    # print("Main loop is now over and we can do other stuff.")

    try:
        gp_convert(input_file, include_size.get())
    except MissingColumnException as ex:
        root = tk.Tk(className=' Task Failed')
        root.geometry('%dx%d+%d+%d' % (w, h, x, y))
        error_message = tk.Label(root, text=str(ex))
        error_message.pack(pady=(30, 0))
        exit_button = tk.Button(root, text='Quit', command=sys.exit)
        exit_button.pack(pady=(15, 0))

        root.mainloop()
    except Exception as ex:
        root = tk.Tk(className=' Task Failed')
        root.geometry('%dx%d+%d+%d' % (w, h, x, y))
        error_message = tk.Label(root, text='Unexpected Error: ' + str(ex))
        error_message.pack(pady=(30, 0))
        exit_button = tk.Button(root, text='Quit', command=sys.exit)
        exit_button.pack(pady=(15, 0))

        root.mainloop()

    # Success Message
    root = tk.Tk(className=' Task Complete')
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    success_message = tk.Label(root, text='GP to ERA Conversion Successful\n\nConverted files will be stored\n in the '
                                          'same directory as the input file.')
    success_message.pack(pady=(45, 0))
    exit_button = tk.Button(root, text='Quit', command=sys.exit)
    exit_button.pack(pady=(15, 0))

    root.mainloop()


def gp_convert(filename, include_size):
    target = os.path.dirname(filename) + '/GP_Import_' + date.today().strftime("%b-%d-%Y") + '/'
    try:
        os.mkdir(target)
    except FileExistsError:
        pass

    drs = []
    dfs_fields = ['Disaster Number', 'Applicant FIPS', 'PW #', 'Project Category', 'Project Title', 'Project Reference',
                  'Is Large', 'Is Withdrawn', 'GP ID', 'GP Process Step', 'GP Last Action', 'GP Last Process Step',
                  'Has 406 Mitigation', 'Approximate Cost', 'CRC Gross', 'CRC Net', 'Fed Share Percentage',
                  'Date Obligated', 'P4 Date', 'MB3 Closeout Step', 'MB3 RFR Step', 'MB3 ID', 'MB3 Sync Date',
                  'GP Sync Date', 'Obligated Versions Count', 'Operational Percent Complete',
                  'Operational Percent Complete As Of Date', 'Project Damage Description', 'Project Scope of Work',
                  'Validation Notes', 'Validation Status', 'SLA Notes', 'SLA Status', 'Recipient Review Notes',
                  'Recipient Review Status', 'Payment Notes', 'Payment Status', 'RFR Notes', 'RFR Status',
                  'EMMIE Upload Notes', 'EMMIE Upload Date', 'SLA Number', 'Payment Total', 'SLA User ID',
                  'Recipient Review User ID', 'Payment User ID', 'RFR User ID', 'Validation User ID',
                  'EMMIE Upload User ID', 'PW Version Review Status', 'GP Project Type', 'Closed Date',
                  'Quarterly Report Date']

    with open(filename, mode='r', encoding='utf-8-sig') as f:
        csv_f = csv.reader(f)
        gp_fields = next(csv_f)
        print('Fields: ', end=''), print(gp_fields)
        if validate_headers(gp_fields):  # Check if all required columns are present
            raise MissingColumnException('Required columns are missing: ' + validate_headers(gp_fields))
        for row in csv_f:
            if row[gp_fields.index('Event')][:4] in drs:
                continue
            else:
                drs.append(row[gp_fields.index('Event')][:4])
    print('DRs: ', end=''), print(drs)

    # Original Method
    for dr in drs:
        with open(filename) as f, open(target + dr + '.csv', 'w', newline='') as dest:
            writer_object = csv.writer(dest)
            writer_object.writerow(dfs_fields)
            csv_f = csv.reader(f)
            next(csv_f)
            for row in csv_f:
                if row[gp_fields.index('Event')][:4] == dr:
                    new_row = transpose(gp_fields, row, include_size)
                    writer_object.writerow(new_row)

    # Standard Method (Use Original Method if alternative doesn't work)
    # for dr in drs:
    #     with open(filename) as f, open(target + dr + '.csv', 'w', newline='') as dest:
    #         writer_object = csv.writer(dest)
    #         writer_object.writerow(dfs_fields)
    #         csv_f = csv.reader(f)
    #         next(csv_f)
    #         for row in csv_f:
    #             if row[gp_fields.index('Event')][:4] == dr:
    #                 new_row = transpose(gp_fields, row, include_size)
    #                 writer_object.writerow(new_row)
    #     df = pd.read_csv(target + dr + '.csv')
    #     df.to_excel(target + dr + '.xlsx', sheet_name=dr, index=False)  # index=True to write row index
    #     print(dr + ' successfully converted to .xlsx')
    #     os.remove(target + dr + '.csv')

    # Alternative Method (using pandas .xlsx direct read/write)
    # for dr in drs:
    #     with open(filename) as f, open(target + dr + '.csv', 'w', newline='') as dest:
    #         writer_object = pd.ExcelWriter(dest, mode='a', engine='openpyxl', datetime_format='DD/MM/YYYY')
    #         writer_object.writerow(dfs_fields)
    #         csv_f = pd.read_csv(f, parse_dates=[0], infer_datetime_format=True)
    #         next(csv_f)
    #         for row in csv_f:
    #             if row[gp_fields.index('Event')][:4] == dr:
    #                 new_row = transpose(gp_fields, row, include_size)
    #                 writer_object.writerow(new_row)
    #     df = pd.read_csv(target + dr + '.csv')
    #     df.to_excel(writer_object, sheet_name=dr, index=False)  # index=True to write row index
    #     print(dr + ' successfully converted to .xlsx')
    #     os.remove(target + dr + '.csv')


def transpose(gp_fields, row, include_size):
    r = [row[gp_fields.index('Event')][:4]]
    x = len(row[gp_fields.index('Subrecipient')]) - row[gp_fields.index('Subrecipient')][::-1].find("(")
    y = len(row[gp_fields.index('Subrecipient')]) - row[gp_fields.index('Subrecipient')][::-1].find(")")
    fips = row[gp_fields.index('Subrecipient')][x:y - 1].strip()
    r.extend([fips, row[gp_fields.index('P/W #')], row[gp_fields.index('Category')], row[gp_fields.index('Title')], ''])
    if include_size:
        is_large = ('y' if row[gp_fields.index('Project Size')] == 'Large' else 'n')
        r.extend([is_large, ''])
    else:
        r.extend(['', ''])
    r.extend([row[gp_fields.index('Project #')], row[gp_fields.index('Process Step')]])
    if row[gp_fields.index('Last Action Date')] != '':
        last_action = datetime.strptime(row[gp_fields.index('Last Action Date')][:-4], '%m/%d/%Y %I:%M %p')
    else:
        last_action = ''
    if row[gp_fields.index('Last Process Step Date')] != '':
        last_process = datetime.strptime(row[gp_fields.index('Last Process Step Date')][:-4], '%m/%d/%Y %I:%M %p')
    else:
        last_process = ''
    r.extend([last_action, last_process])
    r.extend([row[gp_fields.index('Has 406 Mitigation?')], row[gp_fields.index('Approx. Cost')],
              row[gp_fields.index('CRC Gross Cost')], row[gp_fields.index('CRC Net Cost')],
              row[gp_fields.index('% Cost Share')]])
    r.extend('' for _ in range(6))
    r.append(date.today())
    r.extend('' for _ in range(26))
    r.append(row[gp_fields.index('Type')])
    return r


def validate_headers(headers):
    required_columns = ['Project #', 'P/W #', 'Category', 'Title', 'Event', 'Subrecipient', 'Type', 'Process Step',
                        'Project Size', 'Has 406 Mitigation?', 'Approx. Cost', 'CRC Gross Cost', 'CRC Net Cost',
                        '% Cost Share', 'Last Action Date', 'Last Process Step Date']
    missing_columns = ''
    for header in required_columns:
        if header not in headers:
            missing_columns += '\n' + header
        else:
            continue
    return missing_columns


def validate(filename):  # Write Validation method with error reporting
    with open(filename, 'rb') as inp, open(filename, 'wb') as out:
        writer = csv.writer(out)
        for row in csv.reader(inp):
            if row[0]:
                writer.writerow(row)

    if True:
        return False
    else:
        True


class MissingColumnException(Exception):
    pass


if __name__ == "__main__":
    main()
