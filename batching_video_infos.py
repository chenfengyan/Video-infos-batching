#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os
import xlwt
from moviepy.editor import VideoFileClip

file_dir = u"D:/abc"  # Your videos path


class FileCheck:

    def __init__(self):
        self.file_dir = file_dir

    def get_filesize(self, filename):
        u"""
        Get current file size
        """
        file_byte = os.path.getsize(filename)
        return file_byte

    def get_file_times(self, filename):
        u"""
        Get current video duration
        """
        clip = VideoFileClip(filename)
        file_time = clip.duration
        file_width = clip.size[0]
        file_height = clip.size[1]
        u"""
        Clear clip buffer otherwise report OSError: [WinError 6] The handle is invalid
        """
        clip.reader.close()
        clip.audio.reader.close_proc()
        return file_time, file_width, file_height

    def get_all_video_file(self):
        u"""
        Get all the video files in the directory
        """
        video_files = []
        all_files = []
        # Iterate over all files under file dir, including subdirectories
        self.iter_files(file_dir, all_files)
        for f in all_files:
            if self.is_video_file(f):
                video_files.append(f)
        return video_files

    # Iterate the directory
    def iter_files(self, root_dir, all_files=[]):
        for root, dirs, files in os.walk(root_dir):
            for file in files:
                file_name = os.path.join(root, file)
                all_files.append(file_name)


    def is_video_file(self, file):
        suffix = os.path.splitext(file)[1]
        # Add the video filename suffix you need
        if suffix == '.mp4' or suffix == '.mkv' or suffix == '.wmv'\
                or suffix == '.avi' or suffix == 'mpg':
            return True

        return False


def main():

    print(u"=============Started and waiting...")
    fc = FileCheck()
    files = fc.get_all_video_file()
    datas = [[u'FileName', u'Size', u'Duration', u'Width', u'Height', u'CompressRatio']]  # two-dimensional array
    for f in files:
        cell = []
        file_path = os.path.join(file_dir, f)
        file_size = fc.get_filesize(file_path)
        file_times, file_width, file_height = fc.get_file_times(file_path)
        print(u"FileName:{filename},Size:{filesize},Duration:{filetimes},Width:{filewidth},Height:{fileheight}"
		      .format(filename=f, filesize=file_size, filetimes=file_times, filewidth=file_width, fileheight=file_height))
        cell.append(f)
        cell.append(file_size)
        cell.append(file_times)
        cell.append(file_width)
        cell.append(file_height)
        file_compress = file_size / file_times * (1920 * 1080 / (file_width * file_height))
        cell.append(file_compress)
        datas.append(cell)

    wb = xlwt.Workbook()  # Create Workbook
    sheet = wb.add_sheet('data')  # add a sheet named data

    # Cell format

    style = 'pattern: pattern solid, fore_colour yellow; '  # set fore color yellow
    style += 'font: bold on; '  # set font bold
    style += 'align: horz center, vert center; '  # set horz and vert center
    header_style = xlwt.easyxf(style)

    row_count = len(datas)
    col_count = len(datas[0])
    for row in range(0, row_count):
        col_count = len(datas[row])
        for col in range(0, col_count):
            if row == 0:  # set the header cell format
                sheet.write(row, col, datas[row][col], header_style)
            else:
                sheet.write(row, col, datas[row][col])
    wb.save(file_dir + "/video.xls")
    print("file_dir: " + file_dir)
    print(u"=============Completed")
    pass


if __name__ == '__main__':
    main()
