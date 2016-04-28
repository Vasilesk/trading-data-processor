#!/usr/bin/python
# -*- coding: utf8 -*-

import re
import xlsxwriter

class txt_data_processor:
    def __init__ (self, filename):
        file_to_fetch_data = open(filename, "r")

        first_line = file_to_fetch_data.readline()

        # getting headers from first line without `\r\n`
        self.file_header_list = re.split(' ', first_line[:-1])

        self.data_list = []

        # fetching data
        for line in file_to_fetch_data:
            line_data = re.split(' ', line[:-1])
            header_index = 0
            fetched_from_line = {}
            for element in line_data:
                fetched_from_line.update({self.file_header_list[header_index]: element})
                header_index += 1
            self.data_list.append(fetched_from_line)

        # adding extra data
        self.file_header_list.append('<RHIGH>')
        self.file_header_list.append('<RLOW>')
        self.file_header_list.append('<RCLOSE>')

        self.data_list_size = 0
        val_rhigh_average = 0
        val_rlow_average = 0

        for element in self.data_list:
            self.data_list_size += 1

            val_high = float(element['<HIGH>'])
            val_low = float(element['<LOW>'])
            val_open = float(element['<OPEN>'])
            val_close = float(element['<CLOSE>'])

            val_rhigh = (val_high - val_open) * 100 / val_open
            val_rhigh_average += val_rhigh

            val_rlow = (val_open - val_low) * 100 / val_open
            val_rlow_average += val_rlow

            val_rclose = (val_close - val_open) * 100 / val_open

            element.update({'<RHIGH>': val_rhigh})
            element.update({'<RLOW>': val_rlow})
            element.update({'<RCLOSE>': val_rclose})

        if self.data_list_size > 0:
            val_rhigh_average /= self.data_list_size
            val_rlow_average /= self.data_list_size

        # print(self.data_list)

    def write_to_xlsx (self, filename):
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet('GAZP')

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        for header_element in self.file_header_list:
            worksheet.write(row, col, header_element)
            col += 1
        row += 1

        for one_str in self.data_list:
            col = 0
            for key in self.file_header_list:
                worksheet.write(row, col, one_str[key])
                col += 1
            row += 1

        workbook.close()

if __name__ == "__main__":
    data = txt_data_processor('input/GAZP.txt')
    data.write_to_xlsx('output/data.xlsx')
