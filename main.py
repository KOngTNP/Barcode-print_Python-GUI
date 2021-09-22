import os
import shutil
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import barcode
from barcode.writer import ImageWriter
import win32api
import PIL
from PIL import Image
from PIL import ImageFont
import xlsxwriter
import sqlite3 as sq
from datetime import datetime

root = tk.Tk()
root.title('To-Do List')
root.geometry("900x550+500+300")

conn = sq.connect('todo.db')
cur = conn.cursor()
cur.execute('create table if not exists tasks (title text)')
font = ImageFont.load_default()

task = []


# ------------------------------- Functions--------------------------------
def addTask():
    word = e1.get()
    if len(word) == 0:
        messagebox.showinfo('Empty Entry', 'Enter task name')
    elif word in task:
        messagebox.showinfo('Duplicate Entry', 'Enter unique task name')
    else:
        task.append(word)
        bargenerate()
        cur.execute('insert into tasks values (?)', (word,))
        listUpdate()
        e1.delete(0, 'end')

def bargenerate():
    global brcode, fpath
    brcode = barcode.get("code128", e1.get(), writer=ImageWriter())

    dummy = os.getcwd() + "\\images"
    if not os.path.exists(dummy):
        os.makedirs(dummy)
    fpath = os.path.join(dummy, e1.get())

    brcode.save(f'{fpath}', {"module_width": 0.35, "module_height": 3, "font_size": 0, "text_distance": 1, "quiet_zone": 3})
    to_be_resized = Image.open(f'{fpath}.png')  # open in a PIL Image object
    newSize = (400, 80)  # new size will be 500 by 300 pixels, for example
    resized = to_be_resized.resize(newSize, resample=PIL.Image.NEAREST)  # you can choose other :resample: values to get different quality/speed results
    resized.save(f'{fpath}.png', font=font)  # save the resized image

def print_all():
    x = datetime.now()
    dummy = os.getcwd() + "\\files"
    if not os.path.exists(dummy):
        os.makedirs(dummy)
    if len(task) != 0:
        for i in task:
            name = e2.get()
            num = e3.get()
            ref = e4.get()

            workbook = xlsxwriter.Workbook(f'files\{i}.xlsx')
            worksheet = workbook.add_worksheet()

            h1_cell_format = workbook.add_format()
            h1_cell_format.set_bold()
            h1_cell_format.set_font_size(24)

            h2_cell_format = workbook.add_format()
            h2_cell_format.set_bold()
            h2_cell_format.set_font_size(11)

            cell_format = workbook.add_format()
            cell_format.set_font_size(11)

            worksheet.set_margins(left=0.0, right=0, top=0, bottom=0)
            worksheet.insert_image('A9', f'images\{i}.png', {'x_scale': 0.8, 'y_scale': 0.7})


            worksheet.write('A2', f'{name}', h1_cell_format)

            worksheet.write('A4', 'Part number:', h2_cell_format)
            worksheet.write('A5', f'{num}')

            worksheet.write('A6', 'Ref code:', h2_cell_format)
            worksheet.write('A7', f'{ref}')

            worksheet.write('C4', 'Date:', h2_cell_format)
            worksheet.write('C5', f'{x.strftime("%A, %d. %b %y")}')

            worksheet.write('C12', f'{i}')

            workbook.close()
            file = (f'files\{i}.xlsx')


            win32api.ShellExecute(0, "print", file, None, ".", 0)

        messagebox.showinfo("", "Printing barcode")
    else:
        messagebox.showinfo("", "Please enter a subject")



def listUpdate():
    clearList()
    for i in task:
        t.insert('end', i)


def delOne():
    try:
        val = t.get(t.curselection())
        if val in task:
            task.remove(val)
            try:
                os.remove(f'files\{val}.xlsx')
            except:
                pass

            os.remove(f'images\{val}.png')

            listUpdate()
            cur.execute('delete from tasks where title = ?', (val,))


    except:
        messagebox.showinfo('Cannot Delete', 'No Task Item Selected')


def deleteAll():
    mb = messagebox.askyesno('Delete All', 'Are you sure?')
    print("hello1")
    i = 0
    if mb == True:
        while (len(task) != 0):
            task.pop()
            print("hello pop")
        try:

            shutil.rmtree(f'files')

            os.makedirs(os.getcwd() + "\\files")
        except:
            pass
        shutil.rmtree(f'images')
        os.makedirs(os.getcwd() + "\\images")

        cur.execute('delete from tasks')
        listUpdate()

def clearList():
    t.delete(0, 'end')


# def bye():
#     print(task)
#     root.destroy()


def retrieveDB():
    while (len(task) != 0):
        task.pop()
    for row in cur.execute('select title from tasks'):
        task.append(row[0])

def showbrcode(val):
    global photo1
    try:
        val = t.get(t.curselection())
        photo1 = PhotoImage(file=f"images\{val}.png")
        imageLabel.config(image=photo1)
        subLabel.config(text=val)
        # print test button
        b6.place(x=437, y=440)
    except:
        pass

def test_print():
    x = datetime.now()
    i = t.get(t.curselection())

    dummy = os.getcwd() + "\\files"
    if not os.path.exists(dummy):
        os.makedirs(dummy)

    name = e2.get()
    num = e3.get()
    ref = e4.get()

    workbook = xlsxwriter.Workbook(f'files\{i}.xlsx')
    worksheet = workbook.add_worksheet()

    h1_cell_format = workbook.add_format()
    h1_cell_format.set_bold()
    h1_cell_format.set_font_size(24)

    h2_cell_format = workbook.add_format()
    h2_cell_format.set_bold()
    h2_cell_format.set_font_size(11)

    cell_format = workbook.add_format()
    cell_format.set_font_size(11)

    worksheet.set_margins(left=0.0, right=0, top=0, bottom=0)
    worksheet.insert_image('A9', f'images\{i}.png', {'x_scale': 0.8, 'y_scale': 0.7})


    worksheet.write('A2', f'{name}', h1_cell_format)

    worksheet.write('A4', 'Part number:', h2_cell_format)
    worksheet.write('A5', f'{num}')

    worksheet.write('A6', 'Ref code:', h2_cell_format)
    worksheet.write('A7', f'{ref}')

    worksheet.write('C4', 'Date:', h2_cell_format)
    worksheet.write('C5', f'{x.strftime("%A, %d. %b %y")}')

    worksheet.write('C12', f'{i}')

    workbook.close()
    file = (f'files\{i}.xlsx')


    win32api.ShellExecute(0, "print", file, None, ".", 0)

    messagebox.showinfo("", "Printing test barcode")






# ------------------------------- Functions--------------------------------
h1 = ttk.Label(root, text='Edit output', font=("bold 20"))

imageLabel = ttk.Label(root)
subLabel = ttk.Label(root, font=("bold 12"))

l1 = ttk.Label(root, text='To-Do List', font=("bold 12"))
l2 = ttk.Label(root, text='Enter task :', font=("bold 12"))
e1 = ttk.Entry(root, width=30, font=(12))
b1 = ttk.Button(root, text='Add task', width=20, command=addTask)

l3 = ttk.Label(root, text='Enter product name :', font=("bold 14"))
e2 = ttk.Entry(root, width=45, font=("bold 15"))

l4 = ttk.Label(root, text='Enter part number :', font=("bold 14"))
e3 = ttk.Entry(root, width=45, font=("bold 15"))

l5 = ttk.Label(root, text='Enter Ref :', font=("bold 14"))
e4 = ttk.Entry(root, width=45, font=("bold 15"))

t = tk.Listbox(root, height=19, width=30, selectmode='SINGLE', font=(11))
t.bind("<<ListboxSelect>>", showbrcode)
b2 = ttk.Button(root, text='Delete', width=20, command=delOne)
b3 = ttk.Button(root, text='Delete all', width=20, command=deleteAll)


# b4 = ttk.Button(root, text='Exit', width=20, command=bye)
b5 = tk.Button(root, text='Print All', width=46, height=2, font=("bold 16"), bg='Green', fg='White', command=print_all)

b6 = ttk.Button(root, text='Print test', width=20, command=test_print)

retrieveDB()
listUpdate()

# --------------------------- Place geometry --------------------------

#
h1.place(x=15, y=20)

# Entry name text
l3.place(x=30, y=70)
# Enter
e2.place(x=30, y=100)

# Entry PN text
l4.place(x=30, y=160)
# Enter
e3.place(x=30, y=190)

# Entry Ref text
l5.place(x=30, y=250)
# Enter
e4.place(x=30, y=280)

imageLabel.place(x=80, y=350)
subLabel.place(x=230, y=440)

# ---------------------

# Entry task text
l2.place(x=580, y=20)
# Enter
e1.place(x=580, y=50)
# add task butoon
b1.place(x=728, y=80)

# delete button
b2.place(x=580, y=510)
# delete all button
b3.place(x=728, y=510)
# Exit button
# b4.place(x=50, y=510)

# print button
b5.place(x=5, y=475)

l1.place(x=580, y=100)
t.place(x=580, y=130)



root.mainloop()

conn.commit()
cur.close()