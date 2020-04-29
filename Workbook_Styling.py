# Version 1.0 / 2020-04-22 Stats proccessing

import os
import tkinter.filedialog as tk

from openpyxl import Workbook
from openpyxl import load_workbook

default_dir = (r"D:\OneDrive\2.Trans\1.Platts\2020-03 Stats")
os.chdir(default_dir)

sum_file = 'Stats_pandas_sum.xlsx'
# 打开结果所在目录：
# os.startfile(default_dir)

# 目的：调整单元格宽度以适应文件，左对齐（数字除外）：


# 直接打开结果文件？
os.startfile(sum_file)