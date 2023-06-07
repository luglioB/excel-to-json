import pandas
import json
import os

INPUT_FOLDER = os.path.join('.', 'input')
OUTPUT_FOLDER = os.path.join('.', 'output')

filename = input(r"Name of the file to convert - no extension: ")
sheet = input(r"Name of the sheet: ")

found = False

for x in os.listdir(INPUT_FOLDER):
    if x[:-5] == filename:
        found = True
        if not x.endswith('.xlsx'):
            raise 'file extension is not valid (must be xlsx)'
        else:
            try:
                excel_data_df = pandas.read_excel(os.path.join(INPUT_FOLDER, x), sheet_name=sheet)
                thisisjson = excel_data_df.to_json(orient='records')
                thisisjson_dict = json.loads(thisisjson)
            except:
                raise 'file content is not valid'

        with open(os.path.join(OUTPUT_FOLDER, f'{filename}.json'), 'w') as json_file:
            json.dump(thisisjson_dict, json_file)

if found == False:
    raise 'file not found in input folder'