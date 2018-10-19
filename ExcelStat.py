#pylint: disable=W0611,E0401,C0413,R0913,W1201,C0103,R0902
import sys
import os
# from string import Template
import logging
# from copy import deepcopy
import datetime
import yaml
# from fuzzywuzzy import fuzz
from openpyxl import Workbook
from openpyxl import load_workbook
import matplotlib


class ExcelStat(object):
    """
    """
    def __init__(self, folder):
        self.log = logging.getLogger(__class__.__name__)
        self.log.setLevel(logging.DEBUG)
        format_string = '%(asctime)s - %(levelname)s - %(name)s - %(message)s'
        formatter = logging.Formatter(format_string)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        self.log.addHandler(console_handler)

        self.end_date = None
        self.folder = folder
        self.sheet = "Tabelle1"
        self.files = [item for item in os.listdir(folder) if not "~" in item]
        self.log.info("Found files: " + str(self.files))

    def get_data(self, xls):
        file_name = os.path.join(self.folder, xls)
        self.log.info("File name: " + file_name)
        self.file_date = datetime.datetime.strptime(\
                xls.replace(" ", "").replace(".xlsx", "")\
                .replace("Einkaufslisteabdem", ""), "%d.%m.%y")
        self.log.info("File date: " + str(self.file_date))

        xls_data = load_workbook(file_name)

        lst_m, lst_v, self.end_date = get_raw_data(xls_data,
                                                   self.sheet,
                                                   self.file_date)
        self.log.debug("Entries for Marius: " + str(len(lst_m)))
        self.log.debug("Entries for Viki: " + str(len(lst_v)))

        data_dict = get_data_dict(self.file_date, lst_m, lst_v)
        return data_dict

    def get_statistics(self, dic):
        total_time = (self.end_date - self.file_date).days
        self.log.info("End date: " + str(self.end_date))
        self.log.info("Absolute time: " + str(total_time) + " days")

        for lst in dic["Marius"].values():
            total_m = sum([val["amount"] for val in lst])
        self.log.info("Total amount (Marius): " + str(total_m))
        for lst in dic["Viki"].values():
            total_v = sum([val["amount"] for val in lst])
        self.log.info("Total amount (Viki): " + str(total_v))

        self.log.info("Total amount (All): " + str(total_m + total_v))
        self.log.info("Amount per day: " + str(round((total_m + total_v)/total_time, 2)))
        print("###############################################################")



def get_raw_data(xls, sheet, file_date):
    data = [[], [], [], []]
    for i, col in enumerate(xls[sheet].iter_cols(max_col=4)):
        for j, cell in enumerate(col):
            data[i].append(cell.value)

    end_date = "Not found..."
    for item in data[0]:
        if isinstance(item, str) and "Bezahlt am " in item:
            end_date = find_date(item, file_date)

    lst_m = declutter_data(data[0], data[1])
    lst_v = declutter_data(data[2], data[3])

    return lst_m, lst_v, end_date

def get_data_dict(date_key, lst_m, lst_v):
    data_dict = {"Marius": {date_key: []},
                 "Viki": {date_key: []}}

    for val, info in lst_m:
        data_dict["Marius"][date_key].append({"amount": val,
                                              "date": find_date(info, date_key),
                                              "info": find_info(info)})
    for val, info in lst_v:
        data_dict["Viki"][date_key].append({"amount": val,
                                            "date": find_date(info, date_key),
                                            "info": find_info(info)})
    return data_dict

def declutter_data(lst1, lst2):
    new_lst = []
    for cell1, cell2 in zip(lst1, lst2):
        if isinstance(cell1, (int, float)):
            new_lst.append((cell1, cell2))
    return new_lst

def find_date(val, date_key):
    date = []
    if val is None:
        return date
    for char in val:
        if char.isdigit() or char == ".":
            date.append(char)
    if date != []:
        date = adjust_date_str(date, date_key)
    return date

def adjust_date_str(date, date_key=None):
    try:
        if date[1] == ".":
            date.insert(0, "0")
        if date[4] == ".":
            date.insert(3, "0")
        if len(date) == 6:
            date.append(str(date_key.year)[2:])
        date = datetime.datetime.strptime("".join(date), "%d.%m.%y")
    except (TypeError, ValueError, IndexError) as err:
        return []
    return date

def find_info(val):
    if val is None:
        return "No Info"
    try:
        new_val = val.split(" - ")[1]
        if new_val[0] == " ":
            return new_val[1:]
        return new_val
    except:
        if not val[0].isdigit():
            return val
        return val
