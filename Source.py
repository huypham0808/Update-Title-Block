import tkinter as tk
from tkinter import filedialog

import os
from pyautocad import Autocad
import win32com.client
import xlwings as xw
from pyautocad import *


def chon_tep_cad():
    # Hành động khi nhấn nút "Chọn"
    filename = filedialog.askopenfilename()
    entry1.delete(0, tk.END)
    entry1.insert(0, filename)


def chon_tep_excel():
    # Hành động khi nhấn nút "Chọn"
    filename = filedialog.askopenfilename()
    entry2.delete(0, tk.END)
    entry2.insert(0, filename)


def thuc_hien_cong_viec():
    # Hành động khi nhấn nút "Thực hiện công việc"
    try:
        path_excel = entry2.get()
        wb = xw.Book(path_excel)
        sht = wb.sheets[0]
        print(sht.name)
        # Lấy các dạng bản vẽ
        first_row = 2
        last_row = sht.range("A" + str(sht.cells.last_cell.row)).end("up").row
        print(last_row)

        list_layout_excel_old = sht.range(
            "B"+str(first_row)+":B"+str(last_row)).value
        list_layout_excel = [x for x in list_layout_excel_old if x is not None]
        list_TENBV1_excel_old = sht.range(
            "C"+str(first_row)+":C"+str(last_row)).value
        list_TENBV1_excel = [x for x in list_TENBV1_excel_old if x is not None]
        list_TENBV2_excel_old = sht.range(
            "D"+str(first_row)+":D"+str(last_row)).value
        list_TENBV2_excel = [x for x in list_TENBV2_excel_old if x is not None]
        list_TLBV_excel_old = sht.range(
            "E"+str(first_row)+":E"+str(last_row)).value
        list_TLBV_excel = [x for x in list_TLBV_excel_old if x is not None]
        list_LXB_excel_old = sht.range(
            "F"+str(first_row)+":F"+str(last_row)).value
        list_LXB_excel = [x for x in list_LXB_excel_old if x is not None]
        list_LCS_excel_old = sht.range(
            "G"+str(first_row)+":G"+str(last_row)).value
        list_LCS_excel = [x for x in list_LCS_excel_old if x is not None]
        list_BVS_excel_old = sht.range(
            "H"+str(first_row)+":H"+str(last_row)).value
        list_BVS_excel = [x for x in list_BVS_excel_old if x is not None]

        print("------------")
        print(list_BVS_excel)
        acad = win32com.client.Dispatch("AutoCAD.Application")

        path_cad = entry1.get()
        print(path_cad)
        if path_cad == None or path_cad == "":
            doc = acad.ActiveDocument
        else:
            doc = acad.Application.Documents.Open(path_cad)
        # doc = acad.ActiveDocument
        layouts = doc.Layouts

        list_layouts = []
        for i in layouts:
            if i.name != "Model":
                list_layouts.append(i)
            if len(list_layouts) == len(list_layout_excel):
                dem = 0
                for j in range(len(list_layout_excel)):
                    for i in range(len(list_layouts)):
                        if list_layouts[i].name.upper().strip() == list_layout_excel[j].upper().strip():
                            list_elemnent_layout = []
                            for element in list_layouts[i].Block:
                                if element.EntityName == "AcDbBlockReference" and element.HasAttributes and element.name == "DP_NDBV":
                                    list_att = element.GetAttributes()
                                    for att in list_att:
                                        para_name = att.TagString
                                        para_value = att.TextString
                                        if att.TagString == "PROJECT_TITLE1":
                                            att.TextString = list_TENBV1_excel[j]
                                        elif att.TagString == "PROJECT_TITLE2":
                                            att.TextString = list_TENBV2_excel[j]
                                        elif att.TagString == "BVS":
                                            att.TextString = list_BVS_excel[j]
                                        elif att.TagString == "TLBV":
                                            att.TextString = list_TLBV_excel[j]
                                        elif att.TagString == "LXB":
                                            att.TextString = list_LXB_excel[j]
                                        elif att.TagString == "LCS":
                                            att.TextString = list_LCS_excel[j]
                                    thong_bao = "-Xong khung tên của layout:" + \
                                        str(list_BVS_excel[j])+"\n"
                                    dem = dem + 1
                                    result_text.insert(tk.END, thong_bao)
                                    print(thong_bao)
                    # print("----------------")
                # thong_bao1 = "-Xong khung tên toàn bộ file\n"
                thong_bao1 = "-Xong khung tên của " + \
                    str(dem)+"/"+str(len(list_layouts))+" layouts"
                result_text.insert(tk.END, "-------------------------------\n")
                result_text.insert(tk.END, thong_bao1)
                print(thong_bao1)
                wb.close()
                # app = wb.apps.active
                # app.quit()
            else:
                thong_bao1 = "Số layouts Cad không khớp với Excel"
                result_text.insert(tk.END, thong_bao1)
    except:
        thong_bao1 = "Đã xảy ra lỗi. Mời bạn thử lại"
        result_text.insert(tk.END, thong_bao1)
        print(thong_bao1)


def reset_du_lieu():
    # Hành động khi nhấn nút "Reset dữ liệu"
    entry1.delete(0, tk.END)
    entry2.delete(0, tk.END)
    result_text.delete(1.0, tk.END)


# Tạo cửa sổ chính
root = tk.Tk()
root.title("Update Khung Tên")
root.configure(bg="#f0f0f0")  # Thiết lập màu nền cho cửa sổ chính

# Tạo ô nhập liệu 1
entry1 = tk.Entry(root, width=50)
entry1.grid(row=0, column=0, pady=10, padx=10, columnspan=2)
entry1.configure(bg="white")  # Thiết lập màu nền cho ô nhập liệu 1

# Tạo ô nhập liệu 2
entry2 = tk.Entry(root, width=50)
entry2.grid(row=1, column=0, pady=10, padx=10, columnspan=2)
entry2.configure(bg="white")  # Thiết lập màu nền cho ô nhập liệu 2


# Tạo nút "Chọn" cho ô nhập liệu 1
button_chon_tep = tk.Button(root, text="Chọn ACad", command=chon_tep_cad, bg="#4caf50",
                            # Thiết lập màu cho nút "Chọn"
                            fg="white", font=("Arial", 10, "bold"))
button_chon_tep.grid(row=0, column=2, pady=0, padx=0)
# Tạo nút "Chọn" cho ô nhập liệu 2
button_chon_tep = tk.Button(root, text="Chọn Excel", command=chon_tep_excel, bg="#4caf50",
                            # Thiết lập màu cho nút "Chọn"
                            fg="white", font=("Arial", 10, "bold"))
button_chon_tep.grid(row=1, column=2, pady=0, padx=0)

# Tạo nút "Thực hiện công việc"
button_thuc_hien = tk.Button(root, text="Update Khung", command=thuc_hien_cong_viec, bg="#2196f3", fg="white", font=(
    "Arial", 10, "bold"))  # Thiết lập màu cho nút "Thực hiện công việc"
button_thuc_hien.grid(row=2, column=0, pady=5, padx=5, sticky="e")

# Tạo nút "Reset dữ liệu"
button_reset = tk.Button(root, text="Reset dữ liệu", command=reset_du_lieu, bg="#f44336",
                         # Thiết lập màu cho nút "Reset dữ liệu"
                         fg="white", font=("Arial", 10, "bold"))
button_reset.grid(row=2, column=1, pady=5, padx=5, sticky="w")

# Tạo ô hiển thị kết quả
result_text = tk.Text(root, height=10, width=50)
result_text.grid(row=3, column=0, columnspan=3, pady=10, padx=10)
result_text.configure(bg="white")  # Thiết lập màu nền cho ô hiển thị kết quả

# Tạo kích thước cửa sổ cố định và không thể co giãn
root.geometry("425x300")
root.resizable(width=False, height=False)

root.mainloop()
