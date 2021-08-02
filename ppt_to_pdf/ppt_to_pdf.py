# -*- coding:utf-8 -*-
# Author: 小灰机大
# CreateAt: 2021-8-2
# convert ppt to pdf
# outputFolder: ./pdf   inputFolder: ./ppt


import comtypes.client
import os

INPUT_FOLDER_NAME = 'ppt'
OUTPUT_FOLDER_NAME = 'pdf'


# 初始化 ppt 对象
def init_ppt_and_folder(folder):
    if os.path.exists(os.path.join(folder, INPUT_FOLDER_NAME)):
        ppt = comtypes.client.CreateObject("Powerpoint.Application")
        ppt.Visible = 1
        # 判断输出文件夹是否存在
        if not os.path.exists(os.path.join(folder, OUTPUT_FOLDER_NAME)):
            os.mkdir(os.path.join(folder,OUTPUT_FOLDER_NAME))
        return ppt
    else:
        # 判断输入文件夹是否存在
        print('没有找到输入文件夹，请将要转换的 ppt 或者 pptx 文件，放入当前文件夹下的 ppt 文件夹内')
        return False


# 将 ppt 转换为 pdf
def ppt_to_pdf(ppt, input_filename, output_filename, format_type=32):
    output_filename = output_filename + os.sep + (input_filename.split('.')[0]).split(os.sep)[-1] + '.pdf'
    print('outFileName: ', output_filename,'starting...')
    if input_filename[-4:] == '.ppt' or input_filename[-4:] == 'pptx':
        deck = ppt.Presentations.Open(input_filename)
        deck.SaveAs(output_filename, format_type)
        deck.close()
        print(output_filename.split(os.sep)[-1], ' switch successfully')
    return


def convert_files_folder(ppt, folder):
    input_folder = os.path.join(folder, INPUT_FOLDER_NAME)
    output_folder = os.path.join(folder, OUTPUT_FOLDER_NAME)
    files = os.listdir(input_folder)
    pptfiles = [file for file in files if file.endswith(('ppt', 'pptx'))]
    for pptfile in pptfiles:
        path = os.path.join(input_folder, pptfile)
        ppt_to_pdf(ppt, path, output_folder)


if __name__ == '__main__':
    folder = os.getcwd()
    ppt = init_ppt_and_folder(folder)
    if ppt:
        convert_files_folder(ppt, folder)
        ppt.Quit()
