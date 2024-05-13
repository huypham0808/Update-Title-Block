from tkinter import *

win = Tk()
win.title("Xuat excel to CAD")
win.geometry('300x300')
#win['bg'] = 'gray'
win.attributes('-topmost', True)

name = Label(win, text='Label 1', font=('Arial', 14), bg='red', fg='white')
name.place(x = 20, y = 100)

win.mainloop()