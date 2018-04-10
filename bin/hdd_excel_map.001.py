import os
import re
import sys
import xlrd
from collections import defaultdict
import pprint

PATH = r"D:\work\python_dev\hdd_excel_map\sampleexcels"

def prepare_inputs(string_input):
    string_input = str(string_input)
    if re.match("text:", string_input, re.I):
        type_string = "text"
    elif re.match("number:", string_input, re.I):
        type_string = "number"
    else:
        type_string = None
    if type_string == "text":
        string_input = (str(string_input).split("text:u")[1]).replace("'", "")
        string_input = string_input.replace("\\\\", "/")
    elif type_string == "number":
        string_input = (str(string_input).split("number:")[1])
    else:
        string_input = None
    return string_input

def breakdown_excel_files(excel_file_path, sheet_breakdown_dict):
    excel_book_open = xlrd.open_workbook(excel_file_path)
    sheet_open = excel_book_open.sheet_by_index(0)
    excel_book_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    for rows in range(sheet_open.nrows):

        sheet_row_value = sheet_open.row(rows)
        drive_data_split = os.path.splitdrive(prepare_inputs(sheet_row_value[0]))
        key_value_data = drive_data_split[-1]
        drive_name = drive_data_split[0]
        if key_value_data == "/" or re.search("/\$", key_value_data) or re.search("/\.", key_value_data) or re.search("(\*\.\*)$", key_value_data):
            continue
        else:
            data_list = []
            for row_items_values in sheet_row_value[1:]:
                to_insert = prepare_inputs(row_items_values)
                if to_insert:
                    data_list.append(to_insert)
            sheet_breakdown_dict[key_value_data][drive_name] = data_list
    return sheet_breakdown_dict

def process_headings(heading_inputs_string):
    heading_input_list = re.split("[\s,\W]", heading_inputs_string.strip())
    heading_input_list = [items.strip() for items in heading_input_list if items.strip() != ""]
    return heading_input_list

def main():
    # path_input = raw_input("Please enter the folder path: ")
    path_input = PATH
    # heading_inputs = raw_input("Please add in the headings (comma separated): ")
    # if heading_inputs == "":
    #     print ("Please enter headings, code will quit...")
    #     return

    if os.path.exists(path_input):
        excel_file_documents = [os.path.join(path_input, items) for items in os.listdir(path_input) if os.path.splitext(items)[-1] in [".xlsx", ".xls"]]
        if len(excel_file_documents) != 0:
            sheet_breakdown_dict = defaultdict(dict)
            for excel_paths in excel_file_documents:
                sheet_breakdown_dict = breakdown_excel_files(excel_paths, sheet_breakdown_dict)
            missing_folders_items = ""
            difference_folder_items = ""
            for folders, drive_value_infomation in sheet_breakdown_dict.items():
                if len(drive_value_infomation.keys()) != len(excel_file_documents):
                    missing_folders_items += "folder - {0} exists only in {1}\n".format(folders, ",".join(drive_value_infomation.keys()))
                    continue
                current_data = []
                for drive_name, drive_info_list in drive_value_infomation.items():
                    # print current_data
                    if current_data:
                        if drive_info_list != current_data:
                            difference_folder_items += "folder - {0} is different in {1}\n".format(folders, drive_name)
                    current_data = drive_info_list
            print missing_folders_items
            print difference_folder_items
        else:
            print ("No excel files found in the folder")
    else:
        print ("Error: Folder path does not exists")


if __name__ == "__main__":
    main()