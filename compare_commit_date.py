# -*- coding: utf-8 -*-
"""
Created on Fri Mar 31 15:48:54 2023

@author: nviegas001
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Aug 25 14:33:34 2022

@author: nviegas001
"""

# Python program to read an excel file

# import openpyxl module
import openpyxl
import time
import os
import shutil
# from pyautogui import *
# import pyautogui
# import pyperclip
# #import keyboard
# import tkinter as tk
# from tkinter import simpledialog
# from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup

# PATH ES LA DIRECCION DONDE SE ENCUENTRA EL EXCEL
# PARENT DIR ES LA DIR DONDE SE VA A EJECUTAR EL SCRIPT Y DONDE SE VAN A CREAR LAS CARPETAS DE LOS CASOS
path = "/Users/nicolasviegas/Documents/full_screenshot.py/examples/excel.xlsx"
# parent_dir = "C:\\Users\\nviegas001\\python-scripts\\"
# download_dir = "C:\\Users\\nviegas001\\Downloads\\"
github_list = []


def code_list_to_analyze():
    wb_obj = openpyxl.load_workbook(path)

    sheet_obj = wb_obj.active

    cell_obj = sheet_obj.cell(row=1, column=1)
    max_col = sheet_obj.max_column
    max_r = sheet_obj.max_row

    for i in range(2, max_r + 1):
        cell_obj = sheet_obj.cell(row=i, column=1)

        link_github = cell_obj.value

        link = str(link_github)

        github_list.append(str(link))


def obtain_date_commit(chrome, link):
    chrome.get(link)
    commit = chrome.find_element(By.XPATH, "//relative-time[@class='no-wrap']").get_attribute("title")

    print(commit)


def open_list_links():
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option('detach', True)

    chrome = webdriver.Chrome(executable_path=r"/Users/nicolasviegas/Documents/webdriver\chromedriver_mac64",
                              options=chrome_options)

    chrome.get(github_list[0])
    chrome.find_element(By.ID, "login_field").send_keys("nicolasviegas")
    chrome.find_element(By.ID, "password").send_keys("EXWfS2J4#@cn")
    chrome.find_element(By.NAME, 'commit').click()

    for i in github_list:
        if i != 'None':
            obtain_date_commit(chrome, i)

    chrome.quit()


def main():
    code_list_to_analyze()

    open_list_links()


if __name__ == "__main__":
    main()
