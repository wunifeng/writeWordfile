import os

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor, Pt
import re


def set_run(run, font_size, bold, color):
    run.font.size = font_size
    run.bold = bold
    run.font.color.rgb = color
def get_all_files(rootdir):
    files = []
    filelist = os.listdir(rootdir)
    for i in range(0, len(filelist)):
        path = os.path.join(rootdir, filelist[i])
        if os.path.isdir(path):
            files.extend(get_all_files(path))
        if os.path.isfile(path):
            if path.endswith(".docx"):
                files.append(path)
    # print(files)
    return files

if __name__ == '__main__':
    file = get_all_files("D:\\1、讲义")[0]
    print(file)
    (path, filename) = os.path.split(file)
    print(path)
    print(filename)
    outputdir = "D:\\output"




    # 保存文件
    # document.save(outputFile)
