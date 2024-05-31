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
            "A"+str(first_row)+":A"+str(last_row)).value
        list_layout_excel = [x for x in list_layout_excel_old if x is not None]

        list_PROJECT_NAME_excel_old = sht.range(
            "C" + str(first_row) + ":C" + str(last_row)).value
        list_PROJECT_NAME_excel = [
            x for x in list_PROJECT_NAME_excel_old if x is not None]

        list_LOCATION_excel_old = sht.range(
            "D" + str(first_row) + ":D" + str(last_row)).value
        list_LOCATION_excel = [
            x for x in list_LOCATION_excel_old if x is not None]

        list_PROJECT_TITLE1_excel_old = sht.range(
            "E"+str(first_row)+":E"+str(last_row)).value
        list_PROJECT_TITLE1_excel = [
            x for x in list_PROJECT_TITLE1_excel_old if x is not None]
        list_SHEET_TITLE_excel_old = sht.range(
            "F"+str(first_row)+":F"+str(last_row)).value
        list_SHEET_TITLE_excel = [
            x for x in list_SHEET_TITLE_excel_old if x is not None]

        list_DES_excel_old = sht.range(
            "G" + str(first_row) + ":G" + str(last_row)).value
        list_DES_excel = [x for x in list_DES_excel_old if x is not None]

        list_DRN_excel_old = sht.range(
            "H" + str(first_row) + ":H" + str(last_row)).value
        list_DRN_excel = [x for x in list_DRN_excel_old if x is not None]

        list_CHK_excel_old = sht.range(
            "I" + str(first_row) + ":I" + str(last_row)).value
        list_CHK_excel = [x for x in list_CHK_excel_old if x is not None]

        list_IND_excel_old = sht.range(
            "J" + str(first_row) + ":J" + str(last_row)).value
        list_IND_excel = [x for x in list_IND_excel_old if x is not None]

        list_CONTRACTOR_excel_old = sht.range(
            "K" + str(first_row) + ":K" + str(last_row)).value
        list_CONTRACTOR_excel = [
            x for x in list_CONTRACTOR_excel_old if x is not None]

        list_JOBNO_excel_old = sht.range(
            "L" + str(first_row) + ":L" + str(last_row)).value
        list_JOBNO_excel = [x for x in list_JOBNO_excel_old if x is not None]
        # REVISON LEVEL 1
        list_REV_LEVEL1_excel_old = sht.range(
            "M"+str(first_row)+":M"+str(last_row)).value
        list_REV_LEVEL1_excel = [
            x for x in list_REV_LEVEL1_excel_old if x is not None]
        list_REV_DATE1_excel_old = sht.range(
            "N"+str(first_row)+":N"+str(last_row)).value
        list_REV_DATE1_excel = [
            x for x in list_REV_DATE1_excel_old if x is not None]

        list_REV_DESC1_excel_old = sht.range(
            "O" + str(first_row) + ":O" + str(last_row)).value
        list_REV_DESC1_excel = [
            x for x in list_REV_DESC1_excel_old if x is not None]

        list_REV_BY1_excel_old = sht.range(
            "P" + str(first_row) + ":P" + str(last_row)).value
        list_REV_BY1_excel = [
            x for x in list_REV_BY1_excel_old if x is not None]
        # REVISON LEVEL 2
        list_REV_LEVEL2_excel_old = sht.range(
            "Q"+str(first_row)+":Q"+str(last_row)).value
        list_REV_LEVEL2_excel = [
            x for x in list_REV_LEVEL2_excel_old if x is not None]
        list_REV_DATE2_excel_old = sht.range(
            "R"+str(first_row)+":R"+str(last_row)).value
        list_REV_DATE2_excel = [
            x for x in list_REV_DATE2_excel_old if x is not None]

        list_REV_DESC2_excel_old = sht.range(
            "S" + str(first_row) + ":S" + str(last_row)).value
        list_REV_DESC2_excel = [
            x for x in list_REV_DESC2_excel_old if x is not None]

        list_REV_BY2_excel_old = sht.range(
            "T" + str(first_row) + ":T" + str(last_row)).value
        list_REV_BY2_excel = [
            x for x in list_REV_BY2_excel_old if x is not None]
        # REVISON LEVEL 3
        list_REV_LEVEL3_excel_old = sht.range(
            "U"+str(first_row)+":U"+str(last_row)).value
        list_REV_LEVEL3_excel = [
            x for x in list_REV_LEVEL3_excel_old if x is not None]
        list_REV_DATE3_excel_old = sht.range(
            "V"+str(first_row)+":V"+str(last_row)).value
        list_REV_DATE3_excel = [
            x for x in list_REV_DATE3_excel_old if x is not None]

        list_REV_DESC3_excel_old = sht.range(
            "W" + str(first_row) + ":W" + str(last_row)).value
        list_REV_DESC3_excel = [
            x for x in list_REV_DESC3_excel_old if x is not None]

        list_REV_BY3_excel_old = sht.range(
            "X" + str(first_row) + ":X" + str(last_row)).value
        list_REV_BY3_excel = [
            x for x in list_REV_BY3_excel_old if x is not None]
        # REVISON LEVEL 4
        list_REV_LEVEL4_excel_old = sht.range(
            "Y"+str(first_row)+":Y"+str(last_row)).value
        list_REV_LEVEL4_excel = [
            x for x in list_REV_LEVEL4_excel_old if x is not None]
        list_REV_DATE4_excel_old = sht.range(
            "Z"+str(first_row)+":Z"+str(last_row)).value
        list_REV_DATE4_excel = [
            x for x in list_REV_DATE4_excel_old if x is not None]

        list_REV_DESC4_excel_old = sht.range(
            "AA" + str(first_row) + ":AA" + str(last_row)).value
        list_REV_DESC4_excel = [
            x for x in list_REV_DESC4_excel_old if x is not None]

        list_REV_BY4_excel_old = sht.range(
            "AB" + str(first_row) + ":AB" + str(last_row)).value
        list_REV_BY4_excel = [
            x for x in list_REV_BY4_excel_old if x is not None]

        print("------------")
        print(list_SHEET_TITLE_excel)
        acad = win32com.client.Dispatch("AutoCAD.Application")

        path_cad = entry1.get()
        print(path_cad)
        if path_cad == None or path_cad == "":
            doc = acad.ActiveDocument
        else:
            doc = acad.Application.Documents.Open(path_cad)
        # doc = acad.ActiveDocument
        layouts = doc.Layouts
        # layouts = SortList(layoutsID)
        list_layouts = []
        for i in layouts:
            if i.name != "Model":
                list_layouts.append(i)
        # list_layouts.sort()
        if len(list_layouts) == len(list_layout_excel):
            dem = 0
            for j in range(len(list_layout_excel)):
                for i in range(len(list_layouts)):
                    if list_layouts[i].name.upper().strip() == list_layout_excel[j].upper().strip():
                        list_elemnent_layout = []
                        for element in list_layouts[i].Block:
                            if element.EntityName == "AcDbBlockReference" and element.HasAttributes and element.name == "STN_TITLE BOX 11x17":
                                list_att = element.GetAttributes()
                                for att in list_att:
                                    para_name = att.TagString
                                    para_value = att.TextString
                                    if att.TagString == "PROJECT_NAME":
                                        att.TextString = list_PROJECT_NAME_excel[j]
                                    elif att.TagString == "PROJECT_LOCATION":
                                        att.TextString = list_LOCATION_excel[j]
                                    elif att.TagString == "PROJECT_TITLE1":
                                        att.TextString = list_PROJECT_TITLE1_excel[j]
                                    elif att.TagString == "SHEET_TITLE":
                                        att.TextString = list_SHEET_TITLE_excel[j]
                                    elif att.TagString == "DESIGN_BY":
                                        att.TextString = list_DES_excel[j]
                                    elif att.TagString == "DRAWING_BY":
                                        att.TextString = list_DRN_excel[j]
                                    elif att.TagString == "APPROVED_BY":
                                        att.TextString = list_CHK_excel[j]
                                    elif att.TagString == "ST_JOB_NO.":
                                        att.TextString = list_IND_excel[j]
                                    elif att.TagString == "CONTRACTOR_NAME":
                                        att.TextString = list_CONTRACTOR_excel[j]
                                    elif att.TagString == "JOB_NO":
                                        att.TextString = list_JOBNO_excel[j]
                                        # Revision level 1
                                    elif att.TagString == "REV_LEVEL1":
                                        att.TextString = list_REV_LEVEL1_excel[j]
                                    elif att.TagString == "REV_DATE1":
                                        att.TextString = list_REV_DATE1_excel[j]
                                    elif att.TagString == "REV_DESC1":
                                        att.TextString = list_REV_DESC1_excel[j]
                                    elif att.TagString == "REV_BY1":
                                        att.TextString = list_REV_BY1_excel[j]
                                        # Revision level 2
                                    elif att.TagString == "REV_LEVEL2":
                                        att.TextString = list_REV_LEVEL2_excel[j]
                                    elif att.TagString == "REV_DATE2":
                                        att.TextString = list_REV_DATE2_excel[j]
                                    elif att.TagString == "REV_DESC2":
                                        att.TextString = list_REV_DESC2_excel[j]
                                    elif att.TagString == "REV_BY2":
                                        att.TextString = list_REV_BY2_excel[j]
                                        # Revision level 3
                                    elif att.TagString == "REV_LEVEL3":
                                        att.TextString = list_REV_LEVEL3_excel[j]
                                    elif att.TagString == "REV_DATE3":
                                        att.TextString = list_REV_DATE3_excel[j]
                                    elif att.TagString == "REV_DESC3":
                                        att.TextString = list_REV_DESC3_excel[j]
                                    elif att.TagString == "REV_BY3":
                                        att.TextString = list_REV_BY3_excel[j]
                                        # Revision level 4
                                    elif att.TagString == "REV_LEVEL4":
                                        att.TextString = list_REV_LEVEL4_excel[j]
                                    elif att.TagString == "REV_DATE4":
                                        att.TextString = list_REV_DATE4_excel[j]
                                    elif att.TagString == "REV_DESC4":
                                        att.TextString = list_REV_DESC4_excel[j]
                                    elif att.TagString == "REV_BY4":
                                        att.TextString = list_REV_BY4_excel[j]
                                thong_bao = "-Updated successfully for layout:" + \
                                    str(list_layout_excel[j])+"\n"
                                dem = dem + 1
                                result_text.insert(tk.END, thong_bao)
                                print(thong_bao)
            # thong_bao1 = "-Xong khung tên toàn bộ file\n"
            thong_bao1 = "-Updated successfully " + \
                str(dem)+"/"+str(len(list_PROJECT_TITLE1_excel))+" layouts\n"
            result_text.insert(tk.END, "-------------------------------\n")
            result_text.insert(tk.END, thong_bao1)
            result_text.configure(state="disabled")
            print(thong_bao1)
            wb.close()
            # app = wb.apps.active
            # app.quit()
        else:
            thong_bao1 = "Number of layouts from AutoCad NOT MATCH with Excel"
            result_text.insert(tk.END, thong_bao1)
    except:
        thong_bao11 = "Something wrong. Please try again!"
        result_text.insert(tk.END, thong_bao11)
        wb.close()
        print(thong_bao11)


def reset_du_lieu():
    # Hành động khi nhấn nút "Reset dữ liệu"
    entry1.delete(0, tk.END)
    entry2.delete(0, tk.END)
    result_text.delete(1.0, tk.END)


def ExportData():
    template = "C:\\Template\\Data - Copy - Copy.xls"
    if not os.path.isfile(template):
        print("Template not found")
        return
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    layouts = doc.Layouts
    # fix sort list
    # layouts.sort()
    excelApp = xw.App(visible=False)
    wb = excelApp.books.open(template)
    sh = wb.sheets[0]
    row = 2

    for layout in layouts:
        if layout.name != "Model":
            # sh.range("A" + str(row)).value = layout.TabName
            sh.range("A" + str(row)).value = layout.name
            sh.range("B" + str(row)).value = layout.Handle
            sh.range("AG" + str(row)).value = layout.TabOrder
            for ele in layout.Block:
                if ele.EntityName == "AcDbBlockReference" and ele.HasAttributes and ele.Name == "STN_TITLE BOX 11x17":
                    listAtt = ele.GetAttributes()
                    for att in listAtt:
                        paraName = att.TagString
                        paraValue = att.TextString
                        if att.TagString == "PROJECT_NAME":
                            sh.range("C" + str(row)).value = att.TextString
                            # sh.range("AC" + str(row)).value = att.WidthFactor
                        elif att.TagString == "PROJECT_LOCATION":
                            sh.range("D" + str(row)).value = att.TextString
                        elif att.TagString == "PROJECT_TITLE1":
                            sh.range("E" + str(row)).value = att.TextString
                        elif att.TagString == "SHEET_TITLE":
                            sh.range("F" + str(row)).value = att.TextString
                        elif att.TagString == "DESIGN_BY":
                            sh.range("G" + str(row)).value = att.TextString
                        elif att.TagString == "DRAWING_BY":
                            sh.range("H" + str(row)).value = att.TextString
                        elif att.TagString == "APPROVED_BY":
                            sh.range("I" + str(row)).value = att.TextString
                        elif att.TagString == "ST_JOB_NO.":
                            sh.range("J" + str(row)).value = att.TextString
                        elif att.TagString == "CONTRACTOR_NAME":
                            sh.range("K" + str(row)).value = att.TextString
                        elif att.TagString == "JOB_NO.":
                            sh.range("L" + str(row)).value = att.TextString
                        elif att.TagString == "REV_LEVEL1":
                            sh.range("M" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DATE1":
                            sh.range("N" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DESC1":
                            sh.range("O" + str(row)).value = att.TextString
                        elif att.TagString == "REV_BY1":
                            sh.range("P" + str(row)).value = att.TextString
                        elif att.TagString == "REV_LEVEL2":
                            sh.range("Q" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DATE2":
                            sh.range("R" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DESC2":
                            sh.range("S" + str(row)).value = att.TextString
                        elif att.TagString == "REV_BY2":
                            sh.range("T" + str(row)).value = att.TextString
                        elif att.TagString == "REV_LEVEL3":
                            sh.range("U" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DATE3":
                            sh.range("V" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DESC3":
                            sh.range("W" + str(row)).value = att.TextString
                        elif att.TagString == "REV_BY3":
                            sh.range("X" + str(row)).value = att.TextString
                        elif att.TagString == "REV_LEVEL4":
                            sh.range("Y" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DATE4":
                            sh.range("Z" + str(row)).value = att.TextString
                        elif att.TagString == "REV_DESC4":
                            sh.range("AA" + str(row)).value = att.TextString
                        elif att.TagString == "REV_BY4":
                            sh.range("AB" + str(row)).value = att.TextString
            row += 1
    wb.save()
    wb.close()
    excelApp.quit()


def SortList(listData):
    custom_order = ["SF0", "SF1", "SF2"]
    sortedData = sorted(listData, key=lambda x: (
        x not in custom_order, custom_order.index(x) if x in custom_order else float('inf')))
    return sortedData


# Tạo cửa sổ chính
root = tk.Tk()
root.title("Update TitleBlock - @Simpson Strong Tie - CSS-2024")
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
button_chon_tep = tk.Button(root, text="Select ACad", command=chon_tep_cad, bg="orange",
                            # Thiết lập màu cho nút "Chọn"
                            fg="white", font=("Arial", 10, "bold"))
button_chon_tep.grid(row=0, column=2, pady=0, padx=10)
# Tạo nút "Chọn" cho ô nhập liệu 2
button_chon_tep = tk.Button(root, text="Select Excel", command=chon_tep_excel, bg="#4caf50",
                            # Thiết lập màu cho nút "Chọn"
                            fg="white", font=("Arial", 10, "bold"))
button_chon_tep.grid(row=1, column=2, pady=0, padx=0)
# Button Export data
buttonExportData = tk.Button(root, text="Export Data", command=ExportData, bg="black", fg="white", font=(
    "Arial", 10, "bold"))
buttonExportData.grid(row=2, column=2, pady=5, padx=5, sticky="e")
# Tạo nút "Thực hiện công việc"
button_thuc_hien = tk.Button(root, text="Update TitleBlock", command=thuc_hien_cong_viec, bg="#2196f3", fg="white", font=(
    "Arial", 10, "bold"))  # Thiết lập màu cho nút "Thực hiện công việc"
button_thuc_hien.grid(row=2, column=0, pady=5, padx=5, sticky="e")

# Tạo nút "Reset dữ liệu"
button_reset = tk.Button(root, text="Reset Data", command=reset_du_lieu, bg="#f44336",
                         # Thiết lập màu cho nút "Reset dữ liệu"
                         fg="white", font=("Arial", 10, "bold"))
button_reset.grid(row=2, column=1, pady=5, padx=5, sticky="w")

# Tạo ô hiển thị kết quả
result_text = tk.Text(root, height=10, width=50)
result_text.grid(row=3, column=0, columnspan=3, pady=10, padx=10)
result_text.configure(bg="white")  # Thiết lập màu nền cho ô hiển thị kết quả


# Tạo kích thước cửa sổ cố định và không thể co giãn
root.geometry("450x300")
root.resizable(width=False, height=False)

root.mainloop()
