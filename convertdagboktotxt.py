#!/usr/bin/env python
# -*- coding: utf-8 -*-
import logging
import ConfigParser
settings = None
try:
    settings = ConfigParser.ConfigParser()
except:
    logging.basicConfig(filename='debug.log',level=logging.DEBUG)
    logging.debug('Failed to import ConfigParser')
if settings:
    settings.read("settings.txt")
    if settings.get('Logging', 'Active').lower() == "true":
        logging.basicConfig(filename='debug.log',level=logging.DEBUG)
        logging.debug('Logging enabled.')
        import sys
        logging.debug("major verion: {0}".format(sys.version_info[0]))

logging.debug('Importing modules...')
import threading
from Tkinter import Label, Tk
from datetime import datetime
import xlrd
import webbrowser
import os
import ctypes
logging.debug('Finished importing modules.')

class DiaryEntry(object):
    def __init__(self, datum, faktiskt_datum, faktisk_start, faktiskt_slut, faktiskt_arbetadeH, vecka, start, slut, timecode_nyckel, isdeb, kvar, h, m, arbetadeH, debiteradeH, aktivitet, tag, beskrivning):
        self.set_year("")
        self.Datum = datum
        self.Faktiskt_datum = faktiskt_datum
        self.Faktisk_start = faktisk_start
        self.Faktiskt_slut = faktiskt_slut
        self.Faktiskt_arbetadeH = faktiskt_arbetadeH
        self.Vecka = vecka.split('.')[0]
        self.Start = start
        self.Slut = slut
        self.Timecode_nyckel = timecode_nyckel
        self.Isdeb = (isdeb.lower() == "x")
        self.Kvar = kvar
        self.H = h
        self.M = m
        self.ArbetadeH = arbetadeH
        self.DebiteradeH = debiteradeH
        self.Aktivitet = aktivitet
        self.Tag = tag
        self.Beskrivning = beskrivning
        self.set_timecode(None)

    def set_year(self, year):
        self.Year = year
    def set_timecode(self, timecode):
        self.Timecode = timecode
    
    def IsValid(self):
        return (self.Datum and self.Vecka and self.Start and self.Slut and self.Timecode_nyckel and self.ArbetadeH and self.Aktivitet)

    def __str__(self):
        raise NotImplementedError('Need to implement this method. The return value should be a string that displays the activity. Either in some importable format in your time reporting system or something easy to read if time is to be reported manually.')

class Timecode(object):
    def __init__(self, key, timecode, task, timecode_type, customer, time_Available, time_Spent, time_Spent_2016, time_Spent_2017):
        self.Key = key
        self.Timecode = timecode
        self.Task = task
        self.Timecode_type = timecode_type
        self.Customer = customer
        self.Time_Available = time_Available
        self.Time_Spent = time_Spent
        self.Time_Spent_2016 = time_Spent_2016
        self.Time_Spent_2017 = time_Spent_2017

    def __str__(self):
        raise NotImplementedError('Need to implement this method. The return value should be a string that displays the timecode. Either in some importable format in your time reporting system or something easy to read if time is to be reported manually. The method can be called in the function: DiaryEntry.__str__().')

    def is_null(self):
        raise NotImplementedError("Need to implement this method. The return value should be a boolean. If it is true, every diary entry using this timecode will be skipped. Return false always to avoid using this functionality.")

def build_DiaryEntry_objects_from_excel_file(excel_filepath):
    logging.debug('Inside build_DiaryEntry_objects_from_excel_file()')
    # Open the excel file using xlrd
    print("Opening workbook...")
    wb = xlrd.open_workbook(excel_filepath)
    
    def get_years_from_sheets():
        years = []
        for sheet in wb.sheets():
            if sheet.name.startswith(settings.get('Diary', 'DiarySheetPrefix')):
                years.append(sheet.name.split(" ")[1])
        return years

    # Go through all sheets in excel workbook and extract all data into lists with a raw format
    ##### Extract Headers
    def get_headers_from_sheet(sheet_name):
        headers = []
        for sheet in wb.sheets():
            if sheet.name == sheet_name:
                for row in range(0, 1):
                    for col in range(sheet.ncols):
                        headers.append(sheet.cell(row,col).value)
        return headers
    ##### Extract Entries
    def get_entries_from_sheet(sheet_name):
        entries = []
        for sheet in wb.sheets():
            if sheet.name == sheet_name:
                for row in range(1, sheet.nrows):
                    values = []
                    for col in range(sheet.ncols):
                        values.append(sheet.cell(row,col).value)
                    entries.append(values)
        return entries
   
    # Get index positions from headers list
    def get_indexes_from_list(list_to_lookup, values):
        indexes = []
        for i in range(0, len(list_to_lookup)):
            if list_to_lookup[i] in values:
                indexes.append(i)
        return indexes

    # Filter the raw data depending on excel cell format and ultimately add it to a new list
    ##### Diary entries
    def get_parsed_diary_entries(diary_entries, diary_headers):
        time_column_indexes = get_indexes_from_list(diary_headers, [settings.get('Diary', 'StartColumn'), settings.get('Diary', 'EndColumn'), settings.get('Diary', 'TotalHoursColumn'), settings.get('Diary', 'ChargedHoursColumn')])
        date_column_indexes = get_indexes_from_list(diary_headers, [settings.get('Diary', 'DateColumn')])
        parsed_diary_entries = []
        for row in diary_entries:
            parsed_values = []
            for i in range(0, len(row)):
                value = row[i]
                if value and i in date_column_indexes: # DATES GO HERE
                    value = str(xlrd.xldate.xldate_as_datetime(value, wb.datemode).date())
                elif value and i in time_column_indexes: # TIME GO HERE
                    value = str(xlrd.xldate.xldate_as_datetime(value, wb.datemode).time())
                elif isinstance(value, float): # FLOATS GO HERE
                    value = str(value)
                value = value.encode('utf-8')
                parsed_values.append(value)
            parsed_diary_entries.append(parsed_values)
        return parsed_diary_entries
    ##### Timecodes
    def get_parsed_timecode_entries(timecode_entries):
        parsed_timecode_entries = []
        for row in timecode_entries:
            parsed_values = []
            for i in range(0, len(row)):
                value = row[i]
                if isinstance(value, float):
                    pass
                else:
                    value = value.encode('utf-8')
                parsed_values.append(value)
            parsed_timecode_entries.append(parsed_values)
        return parsed_timecode_entries

    # Build DiaryEntry objects with attached Timecode objects
    def build_DiaryEntry_objects(parsed_diary_entries_from_excel, parsed_timeCodes_from_excel):
        diaryEntries = []
        ##### Create DiaryEntry objects using the filtered data
        for parsed_diary_entry_from_excel in parsed_diary_entries_from_excel:
            diaryEntries.append(DiaryEntry(*parsed_diary_entry_from_excel))
        ##### Create Timecode objects using the filtered data
        timecodes = []
        for parsed_timeCode_from_excel in parsed_timeCodes_from_excel:
            timecodes.append(Timecode(*parsed_timeCode_from_excel))
        ##### Attach the timecode objects to the DiaryEntry objects
        for diaryEntry in diaryEntries:
            for timecode in timecodes:
                if timecode.Key == diaryEntry.Timecode_nyckel:
                    if not timecode.is_null():
                        diaryEntry.set_timecode(timecode)
                    break
        return diaryEntries

    diary_entries_for_all_years = {}
    for year in get_years_from_sheets():
        sheet_name_for_year = "{0} {1}".format(settings.get('Diary', 'DiarySheetPrefix'), year)
        diary_entries_for_all_years[year] = build_DiaryEntry_objects(
            get_parsed_diary_entries(
                get_entries_from_sheet(sheet_name_for_year), 
                get_headers_from_sheet(sheet_name_for_year)
            ),
            get_parsed_timecode_entries(
                get_entries_from_sheet(settings.get('Diary', 'TimecodeSheetName')) 
            )
        )
        if diary_entries_for_all_years[year] and not os.path.exists(year):
            os.makedirs(year)

    return diary_entries_for_all_years

def write_txt_file_foreach_week_in_diary(diary_entries):
    logging.debug('Inside write_csv_file_foreach_week_in_diary()')
    def write_txt_file(items, filename):
        # Create the diary content
        diary_content = ""
        for current_weeks_item in items:
            if current_weeks_item.IsValid():
                diary_content += str(current_weeks_item)
                diary_content += "\n"
        diary_content = diary_content.rstrip('\n')
        print("Content being applied to csv template:\n{0}".format(diary_content))
        # Get the template for the file and inject the diary content
        raise NotImplementedError("Check the file_content string below and apply how the file for each week should be printed to the each file.")
        file_content = (
            # TODO: INSERT Beginning file content here
            "{0}\n"
            # TODO: INSERT Ending file content here
            ).format(diary_content)
        # Write content to file
        print("Writing...")
        with open(filename, "w") as text_file:
            text_file.write(file_content)
        print("File written!")
    
    for year_key in diary_entries:
        ### Get weeks
        weeks = []
        for diary_entry in diary_entries[year_key]:
            if diary_entry.Vecka and diary_entry.Vecka not in weeks:
                weeks.append(diary_entry.Vecka)
        ### Get entries foreach week
        for week in weeks:
            a_month = None
            a_entries_for_week = []
            b_month = None
            b_entries_for_week = []
            for diary_entry in diary_entries[year_key]:
                if diary_entry.Vecka is not None and diary_entry.Vecka == week:
                    month = diary_entry.Datum.split('-')[1]
                    if not a_month or a_month == month:
                        a_month = month
                        a_entries_for_week.append(diary_entry)
                    elif not b_month or b_month == month:
                        b_month = month
                        b_entries_for_week.append(diary_entry)
                    else:
                        raise Exception("Unexpected error when parsing months from date string")
            if not b_month:        
                write_txt_file(a_entries_for_week, "{0}/{1}-{2}.txt".format(year_key, settings.get('Export', 'FileName'), week))
            else:
                write_txt_file(a_entries_for_week, "{0}/{1}-{2}A.txt".format(year_key, settings.get('Export', 'FileName'), week))
                write_txt_file(b_entries_for_week, "{0}/{1}-{2}B.txt".format(year_key, settings.get('Export', 'FileName'), week))

root = Tk()
def worker():
    logging.debug('Running worker...')
    prompt = settings.get('ProcessingPrompt', 'Message')
    label1 = Label(root, text=prompt, width=len(prompt))
    label1.pack()
    root.mainloop()

def main():
    threading.Thread(target=worker).start()
    ### 1 Open excel file. 
    ### 2 Extract excel data to the model (The DiaryEntry and Timecode classes).
    ### 3 Compile csv files into the folder for the current year.
    ##### Create one file for each week of the current year and let the file data 
    ##### adhere to the specific format that the company time reporting system needs for them to be imported.
    write_txt_file_foreach_week_in_diary(build_DiaryEntry_objects_from_excel_file(settings.get('Diary', 'DiaryFilename')))
    try:
        logging.debug('Destroying worker...')
        root.destroy()
    except:
        pass
    ### 4. Prompt the user to open Time reporting url in Internet explorer.
    logging.debug('Displaying prompt')
    if settings.get('OpenTimeReportingUrl', 'Active').lower() == "true":
        logging.debug('Prompt is active.')
        if ctypes.windll.user32.MessageBoxA(0, settings.get('OpenTimeReportingUrl', 'PromptMessage'), settings.get('OpenTimeReportingUrl', 'PromptTitle'), 4) == 6:
            webbrowser.get(webbrowser.iexplore).open(settings.get('OpenTimeReportingUrl', 'TimeReportingUrl'))

if __name__ == "__main__":
    main()
