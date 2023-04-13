# -*- coding: utf-8 -*-
"""
Created on Fri Mar 31 15:48:54 2023

@author: nviegas001
"""
import time

import openpyxl
from openpyxl import load_workbook
import pandas as pd
import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

path = "/Users/nicolasviegas/Documents/full_screenshot.py/examples/excel.xlsx"

github_list = []
date_list = []

root = tk.Tk()
root.withdraw()
#path = simpledialog.askstring(title="Compare commits", prompt="Ingrese la direccion del excel, debe terminar en .xlsx")

root = tk.Tk()
root.withdraw()
period = simpledialog.askstring(title="Compare commits",
                                prompt="Ingrese el periodo en el que se valida la fecha (I / U)")

root = tk.Tk()
root.withdraw()
pw = simpledialog.askstring(title="Compare commits",
                                prompt="Ingrese password")


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


def date_list_add(commit_date):
    individual_date = str(commit_date)
    date_list.append(str(individual_date))


def obtain_date_commit(chrome, link):
    chrome.get(link)
    commit_date = chrome.find_element(By.XPATH, "//relative-time[@class='no-wrap']").get_attribute("title")

    date_list_add(commit_date)


def open_list_links():
    chrome_options = webdriver.ChromeOptions()
    #chrome_options.add_argument('--headless') #No funciona en 2do plano
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option('detach', True)

    chrome = webdriver.Chrome(executable_path=r"/Users/nicolasviegas/Documents/webdriver\chromedriver_mac64",
                              options=chrome_options)

    chrome.get(github_list[0])

    chrome.find_element(By.ID, "login_field").send_keys("nicolasviegas")
    chrome.find_element(By.ID, "password").send_keys(pw)
    chrome.find_element(By.NAME, 'commit').click()
    time.sleep(40)

    for i in github_list:
        if i != 'None':
            obtain_date_commit(chrome, i)

    chrome.quit()


def condition_write(column_number_condition):

    workbook = openpyxl.load_workbook(path)
    worksheet = workbook.worksheets[0]
    worksheet.insert_cols(column_number_condition)
    y = 2

    cell_title = worksheet.cell(row=1, column=column_number_condition)
    if column_number_condition == 2:
        cell_title.value = 'Fecha WT'
    else:
        cell_title.value = 'Fecha Update'

    for x in range(len(date_list)):
        cell_to_write = worksheet.cell(row=y, column=column_number_condition)
        cell_to_write.value = date_list[x]
        print(date_list[x])
        y += 1

    workbook.save(path)


def write_excel_file():
    wb_obj = openpyxl.load_workbook(path)

    sheet_obj = wb_obj.active

    if period == 'I' or period == 'i':
        condition_write(2)

    elif period == 'U' or period == 'u':
        condition_write(3)

    else:
        print("Ingrese una periodo valido ( I | U)")
        quit()


def request_validation():
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook.worksheets[0]
    worksheet.insert_cols(4)
    cell_title = worksheet.cell(row=1, column=4)
    cell_title.value = 'Requiere validacion'
    workbook.save(path)

    wb_obj = openpyxl.load_workbook(path)

    sheet_obj = wb_obj.active

    cell_obj = sheet_obj.cell(row=1, column=2)
    max_col = sheet_obj.max_column
    max_r = sheet_obj.max_row

    for i in range(2, max_r + 1):
        cell_obj = sheet_obj.cell(row=i, column=2)

        date_wt = cell_obj.value
        cell_obj_u = sheet_obj.cell(row=i, column=3)
        date_u = cell_obj_u.value
        if date_wt != date_u:
            print("Requiere validacion")
            workbook = openpyxl.load_workbook(path)
            worksheet = workbook.worksheets[0]
            cell_to_write = worksheet.cell(row=i, column=4)
            cell_to_write.value = "yes"
            workbook.save(path)

        else:
            print("No requiere validacion")


def main():

    code_list_to_analyze()

    open_list_links()

    write_excel_file()

    request_validation()


if __name__ == "__main__":
    main()
