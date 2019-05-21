#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os
import xlwt
from moviepy.editor import VideoFileClip

file_dir = u"I:/b_s\H-Legend"  # Your videos path


class FileCheck:

    def __init__(self):
        self.file_dir = file_dir

    def get_filesize(self, filename):
        u"""
        获取文件大小（M: 兆）
        """
        file_byte = os.path.getsize(filename)
        return file_byte

    def get_file_times(self, filename):
        u"""
        获取视频时长（s:秒）
        """
        clip = VideoFileClip(filename)
        # print('video info: ' + str(clip.size[0]))
        file_time = clip.duration  # 时长
        file_width = clip.size[0]  # 宽
        file_height = clip.size[1]  # 高
        u"""
        清除clip缓冲 否则报 OSError: [WinError 6] The handle is invalid句柄无效
        """
        clip.reader.close()
        clip.audio.reader.close_proc()
        return file_time, file_width, file_height
    '''
        def sizeConvert(self, size):  # 单位换算
        K, M, G = 1024, 1024 ** 2, 1024 ** 3
        if size >= G:
            return str(size / G) + 'G Bytes'
        elif size >= M:
            return str(size / M) + 'M Bytes'
        elif size >= K:
            return str(size / K) + 'K Bytes'
        else:
            return str(size) + 'Bytes'

    def timeConvert(self, size):  # 单位换算
        M, H = 60, 60 ** 2
        if size < M:
            return str(size) + u'秒'
        if size < H:
            return u'%s分钟%s秒' % (int(size / M), int(size % M))
        else:
            hour = int(size / H)
            mine = int(size % H / M)
            second = int(size % H % M)
            tim_srt = u'%s小时%s分钟%s秒' % (hour, mine, second)
            return tim_srt
    '''

    def get_all_video_file(self):
        u"""
        获取目录下所有的视频文件
        """
        video_files = []
        all_files = []
        # 遍历file_dir下所有文件，包括子目录
        self.iter_files(file_dir, all_files)
        for f in all_files:
            if self.is_video_file(f):
                video_files.append(f)
        return video_files  # 当前路径下所有非目录子文件

    # 遍历文件夹
    def iter_files(self, root_dir, all_files=[]):
        # 遍历根目录
        for root, dirs, files in os.walk(root_dir):
            for file in files:
                if 'giveup' in root:
                    print('include give up in this file. root:' + root)
                    continue
                file_name = os.path.join(root, file)
                all_files.append(file_name)
            # 不知为何不用，递归调用自身,都可以取到子目录（用分区根目录时）
            # for dirname in dirs:
                # self.iter_files(dirname)

    def is_video_file(self, file):
        suffix = os.path.splitext(file)[1]
        if suffix == '.mp4' or suffix == '.mkv' or suffix == '.wmv'\
                or suffix == '.avi' or suffix == 'mpg':
            return True

        return False


def main():
    print(u"=============开始,文件较多，请耐心等待...")
    fc = FileCheck()
    files = fc.get_all_video_file()
    datas = [[u'文件名称', u'文件大小', u'视频时长', u'视频宽度', u'视频高度', u'压缩率']]  # 二维数组
    for f in files:
        cell = []
        file_path = os.path.join(file_dir, f)
        file_size = fc.get_filesize(file_path)
        file_times, file_width, file_height = fc.get_file_times(file_path)
        print(u"文件名：{filename},大小：{filesize},时长：{filetimes},宽：{filewidth},高：{fileheight}"
              .format(filename=f, filesize=file_size, filetimes=file_times, filewidth=file_width, fileheight=file_height))
        cell.append(f)
        cell.append(file_size)
        cell.append(file_times)
        cell.append(file_width)
        cell.append(file_height)
        file_compress = file_size / file_times * (1920 * 1080 / (file_width * file_height))
        cell.append(file_compress)
        datas.append(cell)

    wb = xlwt.Workbook()  # 创建工作簿
    sheet = wb.add_sheet('data')  # sheet的名称为test

    # 单元格的格式

    style = 'pattern: pattern solid, fore_colour yellow; '  # 背景颜色为黄色
    style += 'font: bold on; '  # 粗体字
    style += 'align: horz centre, vert center; '  # 居中
    header_style = xlwt.easyxf(style)

    row_count = len(datas)
    col_count = len(datas[0])
    for row in range(0, row_count):
        col_count = len(datas[row])
        for col in range(0, col_count):
            if row == 0:  # 设置表头单元格的格式
                sheet.write(row, col, datas[row][col], header_style)
            else:
                sheet.write(row, col, datas[row][col])
    wb.save(file_dir + "/video.xls")  # 文件后缀须为xls。xlsx不兼容。文件直接改后缀即可
    print("file_dir: " + file_dir)
    print(u"=============完成")
    pass


if __name__ == '__main__':
    main()
