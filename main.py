import json
from json_parser import json_to_excel

source_file = 'Source File Path But in .txt format containing JSON data.'
destination_file_location = 'Destinatio File Path in xlsx path.'
main_sheet_name = 'Main Sheet Name which will derive the names of other sheets.'
with open(source_file) as f:
    json_file_object = json.load(f)
    json_to_excel(json_file_object=json_file_object, destination_file_location=destination_file_location,
                  main_sheet_name=main_sheet_name)
