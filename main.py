# Kod wyświetla prostą aplikację okienkową
# Po wybraniu wszystkich opcji program pobiera plik excel ze strony, zmienia jego nazwę, kopiuje jeden z środkowych
# wierszy i dodaje skopiowane wartości to już istniejącego pliku.
# przycisk Start aktywuje się gdy wszystkie 3 przyciski zostaną wciśniętę.
import threading
import tkinter.filedialog
from functools import partial
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import os
import pandas as pd
import csv
import os.path
import time
import requests
import re
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd

url = 'www.example.com/login'
url2 = 'www.example.com/login'
url3 = 'www.example.com/login'
url4 = 'www.example.com/login'
url5 = 'www.example.com/login'

filedialog1 = ''  # folder pobrane
filedialog2 = ''  # csv
start_req = ''
login1 = ''
haslo1 = ''


class Test_csv():
    def __init__(self):
        self.s = Service('C:\webdriver\chromedriver.exe')
        self.browser = webdriver.Chrome(service=self.s)  ###, options=self.options desired_capabilities=caps,

    def login_proces(self):
        self.browser.get(url)
        self.browser.implicitly_wait(2)
        m = self.browser.find_element("id", 'userID')
        m.send_keys(login1)
        time.sleep(0.1)
        m1 = self.browser.find_element("id", 'password')
        m1.send_keys(haslo1)
        time.sleep(0.1)
        m1.send_keys(Keys.TAB)
        m1.send_keys(Keys.ENTER)
        self.browser.implicitly_wait(2)
        self.browser.find_element(By.ID, "CybotCookiebotDialogBodyButtonAccept").click()
        self.browser.implicitly_wait(2)

    def loading_page(self, elem):
        time.sleep(1)
        self.browser.get(url4)
        self.browser.implicitly_wait(1)
        self.browser.get(url5)  # przechodzi do strony bez koniecznosci otwierania do nowego oknas
        m5 = self.browser.find_element("name", 'partNO')
        self.browser.find_element("name", 'partNO').clear()
        m5.send_keys(elem)
        self.browser.find_element("name", "Display").click()
        self.browser.find_element("name", "Download Results").click()

    def download_file_and_rename(self):
        for i in range(0, 9):
            file_exists = os.path.exists(fr'{filedialog1}\NazwaExcela.xls')
            if not file_exists:
                time.sleep(1)
            else:
                break
        time.sleep(1)
        old_name = fr'{filedialog1}\NazwaExcela.xls'
        new_name = fr'{filedialog1}\123456789.csv'
        # Renaming the file
        try:
            os.rename(old_name, new_name)
        except:
            pass

    def csv_files(self, elem):
        with open(fr'{filedialog1}\123456789.csv', 'r') as file:
            reader = csv.reader(file, delimiter='\t')

            items_in_1 = list()
            next(reader)
            items_in_1.append(elem)
            items_in_1.append("&")
            for row in reader:
                items_in_1.append(row[2])

            # print(items_in_1)

            with open(fr'{filedialog1}\gotowe.csv', 'a') as f:  ## tutaj dodac
                writer = csv.writer(f)

                writer.writerow(items_in_1)

        os.remove(fr'{filedialog1}\123456789.csv')  ##tutaj dodać


root = Tk()
frm = ttk.Frame(root, padding=10)
frm.grid()


def filedialog12():
    global filedialog2, start_req
    root.filed = fd.askopenfilename(initialdir="/", title="Select file",
                                    filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
    filedialog2 = root.filed
    start_req += 'A'
    print(filedialog2)
    print(start_req)
    AAA = start_req.find('A')
    BBB = start_req.find('B')
    CCC = start_req.find('C')
    if AAA >= 0 and BBB >= 0 and CCC >= 0:
        start_normal()
    label1['text'] = filedialog2
    label1['background'] = '#9CFF2E'
    return


ttk.Button(frm, text="plik z ktorego bedzie pobierane", command=filedialog12).grid(column=1, row=0)
label1 = ttk.Label(frm, text="podaj scieżkę pliku csv", background='pink')
label1.grid(column=1, row=1, columnspan=2)
label2 = ttk.Label(frm, text="podaj scieżkę pliku Pobrane", background='pink')
label2.grid(column=1, row=2, columnspan=2)


def filedialog():
    global filedialog1, start_req
    root.filed = fd.askdirectory(initialdir="/", title="Select folder")
    filedialog1 = root.filed
    start_req += 'B'
    print(filedialog1)
    print(start_req)
    AAA = start_req.find('A')
    BBB = start_req.find('B')
    CCC = start_req.find('C')
    if AAA >= 0 and BBB >= 0 and CCC >= 0:
        start_normal()
    label2['text'] = filedialog1
    label2['background'] = '#9CFF2E'
    return


ttk.Button(frm, text="folder Pobrane", command=filedialog).grid(column=2, row=0)


def start():
    test1 = Test_csv()
    test1.login_proces()
    with open(fr'{filedialog2}', 'r') as fil:  ## stąd pobieramy dane. fr"{gg}"
        reader = csv.reader(fil, delimiter=',')
        lista = list()
        x = 0
        for row in reader:
            x += 1

            lista.append(row[0])
    print(x)
    y = 0
    for item in lista:
        y += 1
        print("jest: ", (y / x) * 100, " %")
        test1.loading_page(item)
        test1.download_file_and_rename()
        try:
            test1.csv_files(item)
        except:
            pass


start_button = ttk.Button(frm, text="Start", command=start, state='disabled')
start_button.grid(column=3, row=0)


def start_normal():
    start_button['state'] = "enable"


username = StringVar()
ttk.Entry(frm, textvariable=username).grid(column=4, row=0)
password = StringVar()
ttk.Entry(frm, textvariable=password).grid(column=4, row=1)


def login():
    global login1, haslo1, start_req
    login1 = username.get()
    haslo1 = password.get()
    start_req += 'C'
    print(login1, haslo1)
    AAA = start_req.find('A')
    BBB = start_req.find('B')
    CCC = start_req.find('C')
    if AAA >= 0 and BBB >= 0 and CCC >= 0:
        start_normal()

    przycisk.configure(bg='#9CFF2E')
    return


przycisk = Button(frm, text='dodaj login', command=login, bg='red')
przycisk.grid(column=4, row=2)

root.mainloop()
