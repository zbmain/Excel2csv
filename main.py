# -*- coding: utf-8 -*-
# @Time    : 2023-09-18 20:23
# @Author  : zbmain

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import subprocess


def selectFile():
    select = select_var.get()

    if select_var.get() == '文件夹':
        filePath = filedialog.askdirectory(initialdir="/", title="选择目录")
        source_path, fileFullName = filePath, os.path.split(filePath)[1]
        export_path = os.path.join(source_path, f"{fileFullName}-输出目录")
    else:
        filePath = filedialog.askopenfilename(initialdir="/", title="选择文件",
                                                 filetypes=[('xlsx', '*.xlsx'), ('xls', '*.xls')])
        source_path, fileFullName = os.path.split(filePath)
        export_path = os.path.join(source_path, f"{fileFullName}-输出目录")

    if filePath == '' or not os.path.exists(filePath):
        return

    operate = messagebox.askquestion('提示', '确定立即转换？')
    if operate == 'no':
        return

    entry.config(state='normal')
    entry.delete(0, 'end')
    entry.insert(0, export_path)
    entry.config(state='readonly')
    os.makedirs(export_path, exist_ok=True)

    if select == "文件夹":
        list_files_and_directories(filePath, source_path, export_path)
    else:
        file_conversion(filePath, source_path, export_path)

    result = messagebox.showinfo('提示', '完成！\n输出目录：\n' + export_path)
    if result == "ok":
        if os.name == 'posix':
            subprocess.Popen(['open', export_path])
        elif os.name == 'nt':
            os.startfile(export_path)
        else:
            print("不支持的操作系统")


def list_files_and_directories(directory, s, e):
    with os.scandir(directory) as entries:
        for entry in entries:
            if entry.is_file():
                file_conversion(os.path.join(directory, entry.name), s, e)
            elif entry.is_dir():
                print(f"目录: {directory} / {entry.name}")
                # 递归遍历子目录
                list_files_and_directories(entry.path, s, e)


def file_conversion(excel_file: str, s: str, e: str):
    if excel_file.endswith((".xlsx", ".xls")):
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        for sheet_name in sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            cur_sheet_name = sep.get() + sheet_name if len(sheet_names) > 1 else ''
            csv_file_name = f'{excel_file.rsplit(".", 1)[0]}{cur_sheet_name}.csv'
            csv_file_name = csv_file_name.replace(s, e)
            os.makedirs(os.path.split(csv_file_name)[0], exist_ok=True)
            # 保存数据到CSV文件
            df.to_csv(csv_file_name, index=False)


def on_radiobutton_toggle():
    choiceBtn.config(text="选择%s" % select_var.get())


def on_validate_input(P):
    if len(P) <= max_characters:
        return True
    else:
        return False


window = tk.Tk()
window.title('Excel -> CSV')

frame = tk.Frame(window)
frame.grid(padx=5, pady=5, sticky=tk.NW)

sepLabel = tk.Label(frame, text='多页表分割符： \t')
sepLabel.grid(row=1, column=0, columnspan=3, sticky=tk.W)
max_characters = 3
validate_input = frame.register(on_validate_input)
sep = tk.Entry(frame, width=3, text='-', validate="key", validatecommand=(validate_input, "%P"))
sep.insert(0, "-")
sep.grid(row=1, column=1, sticky=tk.W)

modeLabel = tk.Label(frame, text='选择文件模式:')
modeLabel.grid(row=0, column=0, sticky=tk.NW)
typeLabel = tk.Label(frame, text=' 支持xlsx/xls文件')
typeLabel.grid(row=0, column=2, sticky=tk.NW)

entry = tk.Entry(frame, width=22, state='readonly')
entry.grid(row=2, column=0, columnspan=2, padx=0)
choiceBtn = tk.Button(frame, text='选择文件', width=6, command=selectFile)
choiceBtn.grid(row=2, column=2, padx=0)

rb_frame = tk.Frame(frame)
options = ["文件", "文件夹"]
select_var = tk.StringVar()
select_var.set(options[0])
for i, option in enumerate(options):
    rb = tk.Radiobutton(rb_frame, text=option, variable=select_var, value=option, command=on_radiobutton_toggle)
    rb.grid(row=0, column=i)
rb_frame.grid(row=0, column=1)

window.mainloop()
