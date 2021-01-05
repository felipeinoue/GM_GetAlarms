from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
from GetAlarms import *


def OpenExcel():
    root.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))


def Dale():
    try:
        response = CreateAlarms(root.filename)
        root.messagebox = messagebox.showinfo(
            title='Info',
            message=response['message']
            )
    except:
        root.messagebox = messagebox.showwarning(
            title='Warning',
            message='Select xlsx first.'
            )


root = Tk()
root.title('GetAlarms')

# select file
btn_openExcel = Button(root, text='Open xlsx', command=OpenExcel, width='30').grid(row=0, column=0, columnspan=2)

# create alarms
btn_start = Button(root, text='Dale', command=Dale, width='30').grid(row=6, column=0, columnspan=2)

# infinite loop for tkinter
root.mainloop()
