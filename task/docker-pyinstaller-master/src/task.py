import tkinter as tk
from tkinter import *
from tkinter import filedialog, ttk
from tkinter import messagebox
import time
import pandas as pd
import os
import sys

# init the windows
win = tk.Tk()
w = 800  # width for the Tk root
h = 650  # height for the Tk root
# get screen width and height
ws = win.winfo_screenwidth()  # width of the screen
hs = win.winfo_screenheight()  # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

win.title("Task01-file arrange")
win.geometry('%dx%d+%d+%d' % (w, h, x, y))
win.resizable(True, True)
# 如果不想讓使用者能調整視窗大小的話就均設為False
bg = tk.Frame(win, bg='black')
bg.pack()


def InSameFileGrouping(file):
    df = pd.read_excel(file)

    select_win = tk.Tk()
    select_win.title("select the group")
    ws = win.winfo_screenwidth()  # width of the screen
    hs = win.winfo_screenheight()  # height of the screen

    w = 400  # width for the Tk root
    h = 300  # height for the Tk root

    # calculate x and y coordinates for the Tk root window

    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    select_win.geometry('%dx%d+%d+%d' % (w, h, x, y))

    list = Listbox(select_win, selectmode="multiple")
    list.pack()

    def select_all():
        list.select_set(0, END)
    Button(select_win, text='select all', command=select_all).pack()

    new = df.groupby(df["Order Group"])

    for x in new.groups:
        list.insert(END, x)

    def confirm():

        master = tk.Tk()
        master.title("What is filenname")
        w = 400  # width for the Tk root
        h = 300  # height for the Tk root
        # get screen width and height
        ws = win.winfo_screenwidth()  # width of the screen
        hs = win.winfo_screenheight()  # height of the screen
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        master.geometry('%dx%d+%d+%d' % (w, h, x, y))
        mylabel = tk.Label(master, text='filename: ')
        mylabel.pack()

        filename = tk.StringVar(master)
        filename.set("ALLresult")
        filename_entry = tk.Entry(master, textvariable=filename)
        filename_entry.pack()

        def send():
            writer = pd.ExcelWriter(
                f'{(os.path.dirname(os.path.realpath(sys.argv[0])))}/{filename.get()}.xlsx')
            final = []
            for i in list.curselection():
                final.append(list.get(i))

            for x in final:
                split = x
                tmp = new.get_group(x)
                tmp.to_excel(writer, sheet_name=f"{split}_result")
                print(f"{x} is finished")
            writer.save()
            messagebox.showinfo(
                title="work DONE",
                message=f"file {filename.get()} output finished"
            )
            time.sleep(5)
            win.quit()

        string_send = tk.StringVar(master)
        string_send.set("OK!!!")
        btn_Send = tk.Button(master, bg='black', fg='blue',
                             textvariable=string_send, command=send)
        btn_Send.pack()

    string_confirm = tk.StringVar(select_win)
    string_confirm.set("OK!!!")
    btn_Confirm = tk.Button(select_win, bg='black', fg='blue',
                            textvariable=string_confirm, command=confirm)
    btn_Confirm.pack()


def InDiffFileGrouping(file):
    df = pd.read_excel(file)
    new = df.groupby(df["Order Group"])

    select_win = tk.Tk()
    select_win.title("select the group")
    ws = win.winfo_screenwidth()  # width of the screen
    hs = win.winfo_screenheight()  # height of the screen
    w = 400  # width for the Tk root
    h = 300  # height for the Tk root
    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    select_win.geometry('%dx%d+%d+%d' % (w, h, x, y))

    list = Listbox(select_win, selectmode="multiple")
    list.pack()

    def select_all():
        list.select_set(0, END)
    Button(select_win, text='select all', command=select_all).pack()

    for x in new.groups:
        list.insert(END, x)

    def confirm():
        final = []
        for i in list.curselection():
            final.append(list.get(i))
        print(final)
        print(type(final))
        for x in final:
            filename = x
            tmp = new.get_group(x)
            tmp.to_excel(
                f"{(os.path.dirname(os.path.realpath(sys.argv[0])))}/{filename}_result.xlsx")
            print(f"{x} is finished")
        messagebox.showinfo(
            title="work DONE",
            message="finish output"
        )
        time.sleep(5)
        win.quit()

    string_confirm = tk.StringVar(select_win)
    string_confirm.set("OK!!!")

    btn_Confirm = tk.Button(select_win, bg='black', fg='blue',
                            textvariable=string_confirm, command=confirm)
    btn_Confirm.pack()


def select():
    string_select.set('selecting...')

    win.filename = filedialog.askopenfilename(
        initialdir=f"{(os.path.dirname(os.path.realpath(sys.argv[0])))}",
        title="Select file",
        filetypes=(("Excel", "*.xlsx"), ("Excel_old", "*.xls"), ("all files", "*.*")))
    # create scrollbars

    # Create a Frame
    frame = Frame(win)
    v = Scrollbar(frame, orient='vertical')
    h = Scrollbar(frame, orient='horizontal')
    frame.pack(pady=50, padx=50)

    # Create a Treeview widget
    tree = ttk.Treeview(frame, xscrollcommand=h.set,
                        yscrollcommand=v.set)
    df = pd.read_excel(win.filename)
    # config in scrollbar in horizental

    h.config(command=tree.xview)
    h.pack(side=BOTTOM, fill=X)

    # config in scrollbar in horizental

    v.config(command=tree.yview)
    v.pack(side=RIGHT, fill=Y)
    print(df)
    string_select.set(f'selected: {win.filename}')

    # Add new data in  widget
    tree["column"] = list(df.columns)

    # style of table
    style = ttk.Style()
    style.configure("Treewview",
                    background="silver",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="silver"
                    )
    style.configure('group',
                    background="yellow",
                    )

    # Change selected color
    style.map('Treeview',
              background=[('selected', 'light blue')]
              )

    for i in list(df.columns):
        if i == "Order Group":
            tree.column(i, width=120, minwidth=100, anchor=W)
        else:
            tree.column(i, width=120, minwidth=100, anchor=W)
    # formating

    tree["show"] = "headings"

    # For Headings iterate over the columns
    for col in tree["column"]:
        tree.heading(col, text=col, anchor=W)

    # Put Data in Rows
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)

    tree.pack()

    OPTIONS = [
        "一個結果檔案",
        "多個結果檔案"
    ]  # etc

    variable = StringVar(win)
    variable.set(OPTIONS[0])  # default value

    w = OptionMenu(win, variable, *OPTIONS)
    w.pack()

    def confirm():

        MsgBox = messagebox.askyesno(
            title='Selected File',
            message=f"確定要將檔案分割為{variable.get()}",
        )
        if MsgBox == True:
            print("name")
            if(variable.get() == "多個結果檔案"):
                InDiffFileGrouping(win.filename)
            if(variable.get() == "一個結果檔案"):
                InSameFileGrouping(win.filename)
        else:
            for i in tree.get_children():
                tree.delete(i)
            tree.destroy()
            frame.destroy()

    string_fun = tk.StringVar()
    string_fun.set('Choose the function')
    btn_comfirm = tk.Button(win, bg='black', fg='blue',
                            textvariable=string_fun, command=confirm)
    btn_comfirm.pack()


# In[ ]:


string_select = tk.StringVar()
string_select.set('Select file')
# Clear all the previous data in tree

btn_select = tk.Button(win, bg='black', fg='blue',
                       textvariable=string_select, command=select)
btn_select.pack()


# [ ] change the file name
# [ ] choosing the select group

# In[ ]:


win.mainloop()
