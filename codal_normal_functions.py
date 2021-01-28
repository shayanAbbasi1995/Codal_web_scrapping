from openpyxl import workbook
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import timeit


def change_numbers_for_date(text):
    text = text.replace('۳۱', '۳۰')
    text = text.replace('۰', '0')
    text = text.replace('۱', '1')
    text = text.replace('۲', '2')
    text = text.replace('۳', '3')
    text = text.replace('۴', '4')
    text = text.replace('۵', '5')
    text = text.replace('۶', '6')
    text = text.replace('۷', '7')
    text = text.replace('۸', '8')
    text = text.replace('۹', '9')
    text = text.replace('/', '-')
    return text


def change_numbers(text):
    text = text.replace('۰', '0')
    text = text.replace('۱', '1')
    text = text.replace('۲', '2')
    text = text.replace('۳', '3')
    text = text.replace('۴', '4')
    text = text.replace('۵', '5')
    text = text.replace('۶', '6')
    text = text.replace('۷', '7')
    text = text.replace('۸', '8')
    text = text.replace('۹', '9')
    return text


def rep_char(text, header):
    if header is False:
        text = text.replace(')', '')
        text = text.replace('(', '-')
    text = text.replace(',', '')
    text = text.replace('\n', '')
    text = text.replace('\u200c', ' ')
    text = text.replace('-زیان', '(زیان)')
    text = text.replace('-زيان', '(زيان)')
    text = text.replace('-خروج', '(خروج)')
    text = text.replace('-کاهش', '(کاهش)')
    text = text.replace('-(کسر)', 'کسر')
    text = text.replace('\xa0', '')
    text = text.replace('\u200f', '')
    text = change_numbers(text)
    return text


def str_to_int_or_float(value):
    if isinstance(value, bool):
        return value
    try:
        return int(value)
    except ValueError:
        try:
            return float(value)
        except ValueError:
            return value


def get_stock_names(loc, column, from_which_stock, to_which_stock):
    wb = openpyxl.load_workbook(loc)
    sheet = wb.active
    names = list(sheet.columns)[column]
    names_list = []
    i = from_which_stock - 1
    while i <= to_which_stock - 1:
        names_list.append(str(names[i].value))
        i += 1
    return names_list



