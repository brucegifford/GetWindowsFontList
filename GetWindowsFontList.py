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


def get_font_attributes():
    fonts = []
    fonts_folder_path = shell.SHGetFolderPath(0,shellcon.CSIDL_FONTS, 0, 0)

    fonts_folder = win32com.client.Dispatch("Shell.Application").Namespace(fonts_folder_path)

    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel32.GetBinaryTypeW.restype = ctypes.wintypes.BOOL

    for i in range(fonts_folder.Items().Count):
        font_item = fonts_folder.Items().Item(i)
        font_name = fonts_folder.Items().Item(i).Name
        font_path = fonts_folder.Items().Item(i).Path

        #font_type = win32api.GetBinaryType(font_path)
        result = ctypes.c_ulong()
        if kernel32.GetBinaryTypeW(font_path, ctypes.byref(result)):
            if result.value == 6:  # SCS_32BIT_BINARY
                font_type = 'TrueType'
            else:
                font_type = 'OpenType'
        else:
            font_type = 'Unknown'

        font_attributes = {
            'Name': font_name,
            'Path': font_path,
            'Type': 'TrueType' if font_type == win32con.SCS_32BIT_BINARY else 'OpenType',
            # You can add more font attributes here if desired
        }
        fonts.append(font_attributes)

    fonts.sort(key=lambda k: k["Name"])

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
                    data_file.write(font["Name"]+'\n')



    except Exception as ex:
        print('Got an exception in '+app_name)
        print(str(ex))
        print(traceback.format_exc())


if __name__ == '__main__':
    Main()