import tkinter as tk
from tkinter import ttk

win = tk.Tk()
win.title("Tinh toan")
win.geometry('300x400') #size form
win['bg'] = 'orange' #background color
win.attributes('-topmost', True) #luon luon hien thi tren top
#Tao label
name = ttk.Label(win, text = "Dong chu vi du")
name.place(x = 30, y = 30)



win.mainloop()