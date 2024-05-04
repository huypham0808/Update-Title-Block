import tkinter as tk
from tkinter import filedialog

import os
from pyautocad import Autocad
import win32com.client
import xlwings as xw
from pyautocad import *


def selectCadFile():
    fileName = filedialog.askopenfilename()
    entry1.delete(0, tk.END)
    entry1.insert(0, fileName)


def selectExcelFile():
    # Hành động khi nhấn nút "Chọn"
    filename = filedialog.askopenfilename()
    entry2.delete(0, tk.END)
    entry2.insert(0, filename)


def thuc_hien_cong_viec():
    try:
        path_exel = entry1.get()
        wb = xw.Book(path_exel)
        sht = wb.sheets[0]
        print(sht.name)
        # Lay cac bang ve
        firstRow = 2
        lastRow = sht.range("A" + str(sht.cells.last_cell.row)).end("up").row
        print(lastRow)

    except:
        warning1 = "Somethings wrong"
        print(warning1)


# Creat user interface
root = tk.Tk()
root.title("Update Title Block - CSS Team")
root.configure(bg="#f0f0f0")

# Tạo ô nhập liệu 1
entry1 = tk.Entry(root, width=50)
entry1.grid(row=0, column=0, pady=10, padx=10, columnspan=2)
entry1.configure(bg="white")  # Thiết lập màu nền cho ô nhập liệu 1

# Tạo ô nhập liệu 2
entry2 = tk.Entry(root, width=50)
entry2.grid(row=1, column=0, pady=10, padx=10, columnspan=2)
entry2.configure(bg="white")  # Thiết lập màu nền cho ô nhập liệu 2

# Tạo nút "Chọn" cho ô nhập liệu 1
button_chon_tep = tk.Button(root, text="Chọn ACad", command=selectCadFile, bg="#4caf50",
                            # Thiết lập màu cho nút "Chọn"
                            fg="white", font=("Arial", 10, "bold"))
button_chon_tep.grid(row=0, column=2, pady=0, padx=0)
# Tạo nút "Chọn" cho ô nhập liệu 2
button_chon_tep = tk.Button(root, text="Chọn Excel", command=selectExcelFile, bg="#4caf50",
                            # Thiết lập màu cho nút "Chọn"
                            fg="white", font=("Arial", 10, "bold"))
button_chon_tep.grid(row=1, column=2, pady=0, padx=0)

# Tạo nút "Thực hiện công việc"
button_thuc_hien = tk.Button(root, text="Update Khung", command=thuc_hien_cong_viec, bg="#2196f3", fg="white", font=(
    "Arial", 10, "bold"))  # Thiết lập màu cho nút "Thực hiện công việc"
button_thuc_hien.grid(row=2, column=0, pady=5, padx=5, sticky="e")

# Tạo nút "Reset dữ liệu"
# button_reset = tk.Button(root, text="Reset dữ liệu", command=reset_du_lieu, bg="#f44336", fg="white", font=("Arial", 10, "bold"))  # Thiết lập màu cho nút "Reset dữ liệu"
# button_reset.grid(row=2, column=1, pady=5, padx=5, sticky="w")

# Tạo ô hiển thị kết quả
result_text = tk.Text(root, height=10, width=50)
result_text.grid(row=3, column=0, columnspan=3, pady=10, padx=10)
result_text.configure(bg="white")  # Thiết lập màu nền cho ô hiển thị kết quả

# Tạo kích thước cửa sổ cố định và không thể co giãn
root.geometry("425x300")
root.resizable(width=False, height=False)


root.mainloop()
