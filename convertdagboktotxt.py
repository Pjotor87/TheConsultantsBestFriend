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
    def __init__(self, datum, vecka, start, slut, timecode_nyckel, aktivitet, arbetadeH):
        self.set_year("")
        self.set_timecode(None)
        
        self.Datum = datum
        self.Vecka = vecka.split('.')[0]
        self.Start = start
        self.Slut = slut
        self.Timecode_nyckel = timecode_nyckel
        self.Aktivitet = aktivitet
        self.ArbetadeH = arbetadeH

    def set_year(self, year):
        self.Year = year
    def set_timecode(self, timecode):
        self.Timecode = timecode
    def set_faktiskt_datum(self, faktiskt_datum, faktisk_start, faktiskt_slut, faktiskt_arbetadeH):
        self.Faktiskt_datum = faktiskt_datum
        self.Faktisk_start = faktisk_start
        self.Faktiskt_slut = faktiskt_slut
        self.Faktiskt_arbetadeH = faktiskt_arbetadeH
    def set_deb(self, isdeb, debiteradeH):
        self.Isdeb = (isdeb.lower() == "x")
        self.DebiteradeH = debiteradeH
    def set_calculations(self, kvar, h, m):
        self.Kvar = kvar
        self.H = h
        self.M = m
    def set_descriptions(self, tag, beskrivning):
        self.Tag = tag
        self.Beskrivning = beskrivning

    def IsValid(self):
        return (self.Datum and self.Vecka and self.Start and self.Slut and self.Timecode_nyckel and self.ArbetadeH and self.Aktivitet)

    def __str__(self):
        raise NotImplementedError('Need to implement this method. The return value should be a string that displays the activity. Either in some importable format in your time reporting system or something easy to read if time is to be reported manually.')

class Timecode(object):
    def __init__(self, key, timecode, task, timecode_type):
        self.Key = key
        self.Timecode = timecode
        self.Task = task
        self.Timecode_type = timecode_type

    def set_customer(self, customer):
        self.Customer = customer
    def set_time(self, time_Available, time_Spent):
        self.Time_Available = time_Available
        self.Time_Spent = time_Spent

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
        for value in values:
            value_found = False
            for i in range(0, len(list_to_lookup)):
                if list_to_lookup[i] == value:
                    value_found = True
                    indexes.append(i)
            if not value_found:
                indexes.append(None)
                logging.debug("Value not found: {0}".format(value))
        return indexes

    # Filter the raw data depending on excel cell format and ultimately add it to a new list
    ##### Diary entries
    def get_parsed_diary_entries(diary_entries, diary_headers):
        time_column_indexes = get_indexes_from_list(diary_headers, [settings.get('DiaryColumns', 'StartColumn'), settings.get('DiaryColumns', 'EndColumn'), settings.get('DiaryColumns', 'TotalHoursColumn'), settings.get('DiaryColumns', 'ChargedHoursColumn')])
        date_column_indexes = get_indexes_from_list(diary_headers, [settings.get('DiaryColumns', 'DateColumn')])
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
    def get_parsed_timecode_entries(timecode_entries, timecode_headers):
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

    # Sort the data
    ##### Diary Entries
    def get_sorted_parsed_diary_entries(diary_entries, diary_headers):
        sort_order_indexes = get_indexes_from_list(
            diary_headers, [
                settings.get('DiaryColumns', 'DateColumn'), 
                settings.get('DiaryColumns', 'WeekColumn'), 
                settings.get('DiaryColumns', 'StartColumn'), 
                settings.get('DiaryColumns', 'EndColumn'), 
                settings.get('DiaryColumns', 'TimecodeKeyColumn'),
                settings.get('DiaryColumns', 'ActivityColumn'),
                settings.get('DiaryColumns', 'TotalHoursColumn'),

                settings.get('DiaryColumns', 'ActualDateColumn'),
                settings.get('DiaryColumns', 'ActualstartColumn'),
                settings.get('DiaryColumns', 'ActualEndColumn'),
                settings.get('DiaryColumns', 'ActualWorkedHoursColumn'),

                settings.get('DiaryColumns', 'IsDebColumn'),
                settings.get('DiaryColumns', 'ChargedHoursColumn'),

                settings.get('DiaryColumns', 'LeftThisWeekColumn'),
                settings.get('DiaryColumns', 'HoursColumn'),
                settings.get('DiaryColumns', 'MinutesColumn'),

                settings.get('DiaryColumns', 'TagColumn'),
                settings.get('DiaryColumns', 'DescriptionColumn')
            ])

        sorted_parsed_diary_entries = []
        for diary_entry in diary_entries:
            sorted_entry = []
            for sort_order_index in sort_order_indexes:
                sorted_entry.append(diary_entry[sort_order_index])
            sorted_parsed_diary_entries.append(sorted_entry)

        return sorted_parsed_diary_entries

    ##### Timecodes
    def get_sorted_parsed_timecode_entries(timecode_entries, timecode_headers):
        sort_order_indexes = get_indexes_from_list(
            timecode_headers, [
                settings.get('TimecodeColumns', 'KeyColumn'),
                settings.get('TimecodeColumns', 'TimecodeColumn'), 
                settings.get('TimecodeColumns', 'TaskColumn'), 
                settings.get('TimecodeColumns', 'TypeColumn'),

                settings.get('TimecodeColumns', 'CustomerColumn'),

                settings.get('TimecodeColumns', 'TimeAvailableColumn'), 
                settings.get('TimecodeColumns', 'TimeSpentColumn')
            ])
        
        sorted_parsed_timecode_entries = []
        for timecode_entry in timecode_entries:
            sorted_entry = []
            for sort_order_index in sort_order_indexes:
                sorted_entry.append(timecode_entry[sort_order_index])
            sorted_parsed_timecode_entries.append(sorted_entry)
        
        return sorted_parsed_timecode_entries

    # Build DiaryEntry objects with attached Timecode objects
    def build_DiaryEntry_objects(parsed_diary_entries_from_excel, parsed_timeCodes_from_excel):
        diaryEntries = []
        ##### Create DiaryEntry objects using the filtered data
        for parsed_diary_entry_from_excel in parsed_diary_entries_from_excel:
            diaryEntry = DiaryEntry(
                parsed_diary_entry_from_excel[0], 
                parsed_diary_entry_from_excel[1], 
                parsed_diary_entry_from_excel[2], 
                parsed_diary_entry_from_excel[3], 
                parsed_diary_entry_from_excel[4], 
                parsed_diary_entry_from_excel[5],
                parsed_diary_entry_from_excel[6]
            )
            diaryEntry.set_faktiskt_datum(
                parsed_diary_entry_from_excel[7], 
                parsed_diary_entry_from_excel[8], 
                parsed_diary_entry_from_excel[9], 
                parsed_diary_entry_from_excel[10]
            )
            diaryEntry.set_deb(
                parsed_diary_entry_from_excel[11], 
                parsed_diary_entry_from_excel[12]
            )
            diaryEntry.set_calculations( 
                parsed_diary_entry_from_excel[13], 
                parsed_diary_entry_from_excel[14], 
                parsed_diary_entry_from_excel[15]
            )
            diaryEntry.set_descriptions(
                parsed_diary_entry_from_excel[16], 
                parsed_diary_entry_from_excel[17]
            )
            diaryEntries.append(diaryEntry)
        ##### Create Timecode objects using the filtered data
        timecodes = []
        for parsed_timeCode_from_excel in parsed_timeCodes_from_excel:
            timecode = Timecode(
                parsed_timeCode_from_excel[0],
                parsed_timeCode_from_excel[1],
                parsed_timeCode_from_excel[2],
                parsed_timeCode_from_excel[3]
            )
            timecode.set_customer(
                parsed_timeCode_from_excel[4]
            )
            timecode.set_time(
                parsed_timeCode_from_excel[5],
                parsed_timeCode_from_excel[6]
            )
            timecodes.append(timecode)
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
        # Get entries from sheet -> parse -> sort
        ##### Diary
        diary_sheet_name_for_year = "{0} {1}".format(settings.get('Diary', 'DiarySheetPrefix'), year)
        diary_entries_from_sheet = get_entries_from_sheet(diary_sheet_name_for_year)
        diary_headers_from_sheet = get_headers_from_sheet(diary_sheet_name_for_year)
        parsed_diary_entries = get_parsed_diary_entries(diary_entries_from_sheet, diary_headers_from_sheet)
        sorted_parsed_diary_entries = get_sorted_parsed_diary_entries(parsed_diary_entries, diary_headers_from_sheet)
        ##### Timecodes
        timecode_sheet_name = settings.get('Diary', 'TimecodeSheetName')
        timecode_entries_from_sheet = get_entries_from_sheet(timecode_sheet_name)
        timecode_headers_from_sheet = get_headers_from_sheet(timecode_sheet_name)
        parsed_timecode_entries = get_parsed_timecode_entries(timecode_entries_from_sheet, timecode_headers_from_sheet)
        sorted_parsed_timecode_entries = get_sorted_parsed_timecode_entries(parsed_timecode_entries, timecode_headers_from_sheet)
        # Build diary entry objects for the year
        diary_entries_for_all_years[year] = build_DiaryEntry_objects(sorted_parsed_diary_entries, sorted_parsed_timecode_entries)
        # Create a directory for the year
        if diary_entries_for_all_years[year] and not os.path.exists(year):
            os.makedirs(year)

    return diary_entries_for_all_years

def write_txt_file_foreach_week_in_diary(diary_entries):
    logging.debug('Inside write_txt_file_foreach_week_in_diary()')
    def write_txt_file(items, filename):
        # Create the diary content
        diary_content = ""
        for current_weeks_item in items:
            if current_weeks_item.IsValid():
                diary_content += str(current_weeks_item)
                diary_content += "\n"
        diary_content = diary_content.rstrip('\n')
        print("Content being applied to txt template:\n{0}".format(diary_content))
        # Get the template for the file and inject the diary content
        file_content = (

            "{0}"

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
    ### 3 Compile txt files into the folder for the current year.
    ##### Create one file for each week of the current year and let the file data 
    ##### adhere to the specific format that the time reporting system needs for them to be imported.
    write_txt_file_foreach_week_in_diary(build_DiaryEntry_objects_from_excel_file(settings.get('Diary', 'DiaryFilename')))
    try:
        logging.debug('Destroying worker...')
        root.destroy()
    except:
        pass
    ### 4. Prompt the user to open MyTime in Internet explorer
    logging.debug('Displaying prompt')
    if settings.get('OpenTimeReportingUrl', 'Active').lower() == "true":
        logging.debug('Prompt is active.')
        if ctypes.windll.user32.MessageBoxA(0, settings.get('OpenTimeReportingUrl', 'PromptMessage'), settings.get('OpenTimeReportingUrl', 'PromptTitle'), 4) == 6:
            webbrowser.get(webbrowser.iexplore).open(settings.get('OpenTimeReportingUrl', 'TimeReportingUrl'))

if __name__ == "__main__":
    main()

