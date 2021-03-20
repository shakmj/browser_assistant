# import library
from tkinter import *
import tkinter as tk
import glob
from functools import partial
from selenium import webdriver
from openpyxl import load_workbook

##
print("hii1")

## global configuration variable.
button_space = 50
#my_profile = webdriver.FirefoxProfile(
#    r"c:\Users\shak\AppData\Roaming\Mozilla\Firefox\Profiles\lsamc65l.default-release")
#driver = webdriver.Firefox(my_profile)
driver = webdriver.Firefox()

print("hii2")


##
# functions
def forward_page(url_index, url_list):
    print("url index " + str(url_index))
    print(url_list[url_index])
    driver.get(url_list[url_index])
    return


def toscreen(category_choiced):
    global secscreen
    secscreen = tk.Tk()
    print(category_choiced)
    global flag_of_state
    flag_of_state = 3
    #global secscreen
    #ategory_choiced1 = category_choiced+ ".xlsx"
    pages, url = openfile(category_choiced)
    number_of_pages = len(pages)
    if number_of_pages < 10:
        M = number_of_pages * button_space
        N = 200
    else:
        M = number_of_pages * button_space
        N = 400
    category_choicedname = category_choiced + "menu"
    obj = init_secsc(category_choicedname, M, N, pages, category_choiced, forward_page, url)
#def init_secsc(screen_name, M, N, choices, label, fun, url):

    return category_choiced


def getfilesname():
    # read list of xls in current folder
    xls_list = glob.glob("*.xlsx")
    number_of_choices = len(xls_list)
    if number_of_choices < 10:
        M = number_of_choices * button_space
        N = 200
    else:
        M = number_of_choices * button_space
        N = 400
    return xls_list, number_of_choices, M, N


def init_mainsc(screen_name, M, N, choices, label, fun):
    # create window:
    button_obj = []
    screen_name = tk.Tk()
    screen_name.wm_title(label)
    # configure geometry
    geo = str(N) + "x" + str(M)
    screen_name.geometry(geo)
    for i, obj in enumerate(choices):
        print(i)
        button_obj.append(tk.Button(screen_name, text=obj, command=partial(fun, obj)))
    # print (button_obj)
    # configure window
    for i, btn in enumerate(button_obj):
        btn.grid(row=i + 1, column=2)
    return screen_name


def init_secsc(screen_name, M, N, choices, label, fun, url):
    # create window:
    global secscreen
    button_obj = []
    #screen_name = tk.Tk()
    secscreen.wm_title(label)
    # configure geometry
    geo = str(N) + "x" + str(M)
    secscreen.geometry(geo)
    for i, obj in enumerate(choices):
        print(i)
        button_obj.append(tk.Button(secscreen, text=obj, command=partial(fun, i, url)))
    # print (button_obj)
    # configure window
    for i, btn in enumerate(button_obj):
        btn.grid(row=i + 1, column=2)
    return secscreen


def openfile(file_choiced):
    workbook = load_workbook(filename=file_choiced)
    sheet = workbook.active
    print(sheet.title)
    sheet_row = sheet.max_row
    print(sheet_row)
    pages = []
    url = []
    for i in range(1, sheet_row + 1, 1):
        # print(i)
        cells_page = sheet.cell(row=i, column=1).value
        cells_url = sheet.cell(row=i, column=2).value
        pages.append(cells_page)
        url.append(cells_url)
    print(pages)
    print(url)
    print("------------------------------------------------------------------------------")
    print("------------------------------------------------------------------------------")
    return pages, url


# make buutons
# by for loop:
# xls_name = xls_list
button_obj_mainsc = []
# print (xls_name[0])
# file_choiced = "industrial.xlsx"
### main loop
flag_of_state = 1
bast_f = 1
print("hii")


categores, numnerOfChoices, m, n = getfilesname()
main_sc = init_mainsc("main_sc", m, n, categores, "main categories", toscreen)

while True:
    while flag_of_state == 1:
        main_sc.update()
        print(str(main_sc.winfo_exists()))
        if str(main_sc.winfo_exists())=="1":
            main_sc.deiconify()
            print("ooo4")


    while flag_of_state == 3:
        print("ooo1")
        try:
            print("ooo2")
            secscreen.update()
            main_sc.withdraw()
        except:
            print("ooo")
            flag_of_state =1


