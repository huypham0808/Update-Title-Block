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















#Creat user interface
root = tk.Tk()
root.title("Update Title Block - CSS Team")
root.configure(bg="#f0f0f0")

entry1 = tk.Entry(root, width=50)
entry1.grid(row=0, column=0, pady=10, padx=10, columnspan=2)
entry1.configure(bg="white")  # Thiết lập màu nền cho ô nhập liệu 1

root.mainloop()