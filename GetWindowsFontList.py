import os
import sys
import json
import argparse
import traceback
import win32api
import win32con
import win32com.client
from win32com.shell import shell, shellcon
import ctypes

"""
font_attribute_indexes = {
    0: 'name',
    1: 'path',
    3: 'file_size',
    4: 'file_type',
    5: 'font_family',
    6: 'font_style',
    7: 'font_version',
    8: 'unique_font_dentifier',
    9: 'description',
    10: 'copyright_notice',
    11: 'trademark',
    12: 'date_installed'
}
"""
font_attribute_indexes = {
    0: 'Title',
    1: 'Font style',
    2: 'Show/hide',
    3: 'Designed for',
    4: 'Category',
    5: 'Designer/foundry',
    6: 'Font Embeddability',
    7: 'Font type',
    8: 'Family',
    9: 'Date created',
    10: 'Date modified',
    11: 'File size',
    12: 'Collection',
    13: 'Font file names',
    14: 'Font version'
}
"""
[
0    'Modern Regular',
1    'Regular',
2    'Show',
3    'Latin',
4    'Text',
5    'Microsoft Corporation',
6    'Installable',
7    'Raster',
8    'Modern',
9    '',
10    '\u200e7/\u200e13/\u200e2023 \u200f\u200e8:52 PM',
11    '8.50 KB',
12    '',
13    'C:\\WINDOWS\\Fonts\\MODERN.FON',
14    '0.00', 
'', '', '', '', '']
"""

def strip_LRM_chars(input_str):
    if input_str:
        output_str = input_str.encode('ascii', 'ignore')
        output_str = output_str.decode()
        return output_str
    else:
        return input_str


def get_font_attributes():
    fonts = []
    fonts_folder_path = shell.SHGetFolderPath(0,shellcon.CSIDL_FONTS, 0, 0)

    fonts_folder = win32com.client.Dispatch("Shell.Application").Namespace(fonts_folder_path)

    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel32.GetBinaryTypeW.restype = ctypes.wintypes.BOOL

    for i in range(fonts_folder.Items().Count):
        font_item = fonts_folder.Items().Item(i)
        font_attributes = {}
        for index, key in font_attribute_indexes.items():
            value = fonts_folder.GetDetailsOf(font_item, index)
            if value == "":
                value = None
            font_attributes[key] = value
        font_attributes['Date modified'] = strip_LRM_chars(font_attributes['Date modified'])
        font_attributes['Date created'] = strip_LRM_chars(font_attributes['Date created'])
        """
        values = []
        for j in range(30):
            value = fonts_folder.GetDetailsOf(font_item, j)
            values.append(value)
        #print(values)
        """
        font_name = fonts_folder.Items().Item(i).Name
        font_path = fonts_folder.Items().Item(i).Path
        font_attributes['path'] = font_path

        #font_type = win32api.GetBinaryType(font_path)
        result = ctypes.c_ulong()
        if kernel32.GetBinaryTypeW(font_path, ctypes.byref(result)):
            if result.value == 6:  # SCS_32BIT_BINARY
                font_type = 'TrueType'
            else:
                font_type = 'OpenType'
        else:
            font_type = 'Unknown'
        font_attributes['font_type'] = font_type

        if 'Font file names' in font_attributes and font_attributes['Font file names'] is not None:
            font_attributes['Font file names'] = font_attributes['Font file names'].replace('\\','/')
        if 'path' in font_attributes:
            font_attributes['path'] = font_attributes['path'].replace('\\','/')
        font_attributes_old = {
            'Name': font_name,
            'Path': font_path,
            'Type': 'TrueType' if font_type == win32con.SCS_32BIT_BINARY else 'OpenType',
            # You can add more font attributes here if desired
        }
        fonts.append(font_attributes)

    fonts.sort(key=lambda k: k["Title"].lower())

    return fonts


def Main():
    app_name = "GetWindowsFontList"

    try:
        good_args = []
        for arg in sys.argv[1:]:
            if (not arg.startswith('#')) and (not arg.startswith('--#')) and (not arg.startswith('-#')):
                good_args.append(arg)

        parser = argparse.ArgumentParser(description='Get Windows Font List', fromfile_prefix_chars='@')
        parser.add_argument('-output', default=None, help="file name for outputting json file font list", required=False)
        parser.add_argument('-font_names', default=None, help="file name for outputting json file font list", required=False)


        args = parser.parse_args(good_args)

        font_list = get_font_attributes()
        for font in font_list:
            print(font)

        if args.output:
            with open(args.output, "w", encoding='utf-8') as data_file:
                data_file.write(json.dumps(font_list, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False))

        if args.font_names:
            with open(args.font_names, "w", encoding='utf-8') as data_file:
                for font in font_list:
                    data_file.write(font["Title"]+'\n')



    except Exception as ex:
        print('Got an exception in '+app_name)
        print(str(ex))
        print(traceback.format_exc())


if __name__ == '__main__':
    Main()