'''
Created by Andy Czeropski
'''

import csv
import json
import pandas as pd
import os
import openpyxl
import PySimpleGUI as sg


def main():
    '''
    Creates a GUI to enable users to convert existing csv/excel files into json

    :input: .xlsx/csv filepath
    :input: .json filepath
    :outputL converted .json file
    '''

    sg.theme('Reddit')

    jsonChoices = ['records', 'split', 'index', 'table', 'None']

    layoutSize = (320, 170)

    layout = [[sg.Text('Enter .csv or .xlsx file to convert to json')],
              [sg.Text('Excel/CSV file:', size=(19, 1)),
               sg.FileBrowse(key='excel_fp')],
              [sg.Text('JSON file name:', size=(19, 1)),
               sg.Input(key='json_name', size=(15, 1))],
              [sg.Text('Folder Name (in Ignition)', size=(19, 1)),
               sg.Input(key='ignition_fn', size=(15, 1))],
              [sg.Text('__________________________', size=(25, 1))],
              [sg.Submit(key='submit', size=(10, 1)), sg.Cancel(key='cancel', size=(10, 1))]]

    window = sg.Window('Convert Excel to JSON', layout, size=layoutSize)

    while True:
        # Read any events that took place, and associated values. YHadded: timeout argument to update windows in defined frequency
        event, values = window.read(timeout=300)

        try:
            if event == 'submit':
                # Check if file extension included in user provided name - if no, add file extension to name
                if len(values['json_name']) > 4:
                    if values['json_name'][-5:] == '.json' or values['json_name'][-4:] == '.JSON':
                        json_fp = os.path.dirname(
                            values['excel_fp']) + '/' + values['json_name']
                    else:
                        json_fp = os.path.dirname(
                            values['excel_fp']) + '/' + values['json_name'] + '.json'
                else:
                    json_fp = os.path.dirname(
                        values['excel_fp']) + '/' + values['json_name'] + '.json'

                if os.path.basename(values['excel_fp'][-5:]) == ".xlsx":
                    df = pd.read_excel(values['excel_fp'])

                elif os.path.basename(values['excel_fp'][-4:]) == ".csv":
                    df = pd.read_csv(values['excel_fp'])

                else:
                    sg.popup(
                        "Incorrect filetype! This program only supports .csv and .xlsx file extensions")

                # Create string from excel df
                df_json = pd.DataFrame.to_json(
                    df, indent=4, orient='records')
                # Create list from strings to input into json file
                json_dict = {'name': "MM", 'tagtype': 'Folder'}

                if values['ignition_fn']:
                    json_dict['name'] = values['ignition_fn']
                json_dict['tags'] = json.loads(df_json)

                with open(json_fp, 'w') as f:
                    json.dump(json_dict, f)

                sg.popup('Successfully created json.')

                # print(f'You clicked {event}')
                # print(f'You chose filenames {values[0]} and {values[1]}')
        except OSError as err:
            print("OS error: {0}".format(err))

        # Allows window to be closed on exit or 'X'
        if event == 'EXIT' or event == sg.WIN_CLOSED or event == 'cancel':
            window.close()
            # exit()
            break


if __name__ == '__main__':
    main()
