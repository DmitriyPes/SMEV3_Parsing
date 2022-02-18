# -*- coding: UTF-8 -*-
import codecs
import configparser
import filecmp
import ast
import sys

import zip_unicode
import os
import tempfile
import io
from pathlib import Path
from selenium import webdriver
import time
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import re
import traceback
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
import pandas
import numpy as np
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import win32com.client
from difflib import SequenceMatcher
import email
import email.mime.application
import mimetypes
import locale
from tempfile import mkdtemp
import zipfile
import ntpath
import datetime
from smtplib import SMTPDataError
import shutil
import docx
import logging
"""
from selenium.webdriver import ActionChains

actions = ActionChains(driver)
actions.move_to_element(element).click().perform()
"""

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
chromedriver = ROOT_DIR + r"\chromedriver"
files_to_send = []
info_comp = []
pr_info_comp = []
test_info_comp = []
Log_Format = "%(levelname)s %(asctime)s - %(message)s"
logging.basicConfig(filename='logfile.log', filemode='a', format=Log_Format, level = logging.INFO)
logger = logging.getLogger()
handler = logging.FileHandler('logfile.log')
logger.addHandler(handler)

class web:

    def __init__(self, link):
        self.options = webdriver.ChromeOptions()
        self.preferences = ''
        self.set_preferencies()
        self.set_options()
        self.link = link
        try:
            self.browser = webdriver.Chrome(executable_path=chromedriver, options=self.options)
        except:
            logger.error("Необходимо обновить версию ChromeDriver. " + str(datetime.datetime.now()))
            logger.info("Аварийное отключение программы. " + str(datetime.datetime.now()))
            sys.exit()
        self.msg = MIMEMultipart()
        self.connection(link)
        self.filter = False
        self.filter_word = ''
        self.test_seg = ""
        self.pr_seg = ""
        self.dest_mails = []
        self.mail_subject = ""
        self.send_addr = ''
        self.send_password = ''
        self.vs_name = ''

    def set_options(self):
        #self.options.add_argument('headless')
        self.options.add_argument("--start-maximized")
        self.options.add_experimental_option("prefs", self.preferences)
        self.options.add_experimental_option('prefs', {
            "download.default_directory": str(ROOT_DIR) + r"\downloads",
            "download.directory_upgrade": True,
            "download.prompt_for_download": False,  # To auto download the file
            "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
        })
        self.options.add_experimental_option('excludeSwitches', ['enable-logging'])

    def set_preferencies(self):
        self.preferences = {"download.default_directory": str(ROOT_DIR) + r"\downloads",
                            "safebrowsing.enabled": "false"}

    def close_connection(self):
        logger.info('Отключение от сайта ' + self.link + ' ' + str(datetime.datetime.now()))
        try:
            self.browser.close()
            logger.info('Отключение прошло успешно. ' + str(datetime.datetime.now()))
        except WebDriverException:
            logger.exception('Произошла ошибка при отключении от сайта ' + self.link + ' ' + str(datetime.datetime.now()))
    def get_html_text(self):
        return self.browser.page_source

    def connection(self, link):
        logger.info("Подключение к сайту " + link + " " + str(datetime.datetime.now()))
        try:
            self.browser.get(link)
            time.sleep(1)
            logger.info("Подключение прошло успешно. " + str(datetime.datetime.now()))
            return self.browser.page_source
        except WebDriverException:
            logger.exception("Возникла ошибка при подключении к сайту " + self.link + ' ' + str(datetime.datetime.now()))
            return ''

    def __del__(self):
        self.close_connection()

def excel_connect(path):
    pandas.set_option('display.max_rows', None)
    pandas.set_option('display.max_columns', None)
    pandas.set_option('display.max_colwidth', None)

    vs_info = pandas.read_excel(path, index_col=None, na_values=['NA'], usecols="F,E,H,I")
    vs_info_dict = vs_info.to_dict('list')
    return vs_info_dict

def rename_last_downloaded_zip(end_let):
    paths = sorted(Path(str(ROOT_DIR) + r"\downloads\\").iterdir(), key=lambda f: f.stat().st_mtime)
    filename = paths[len(paths)-1]
    try:
        os.rename(filename, str(filename).replace('.zip', end_let))
    except FileExistsError:
        os.remove(filename)

def get_last_downloaded_file():
    paths = sorted(Path(str(ROOT_DIR) + r"\downloads\\").iterdir(), key=lambda f: f.stat().st_mtime)
    filename = paths[len(paths) - 1]
    return filename

def open_tab(html, link, mode):
    try:
        html.browser.execute_script("window.open();")
        main_window = html.browser.current_window_handle
        html.browser.switch_to.window(html.browser.window_handles[1])
        html.browser.get(link)
        time.sleep(0.1)
    except:
        main_window = ''
    end_let = ''
    if mode == 'TestS':
        end_let = '_test_seg.zip'
    if mode == 'Product':
        end_let = '_productive_seg.zip'
    if end_let != '':
        pass
    if mode == 'Name':
        return name

def make_dir(name):
    try:
        os.mkdir(name)
    except:
        pass

def check_vs(html, vs_info, name):
    tmp = (name.replace('<h3> Новость содержит информацию о следующем ВС - ', '')).replace(' </h3>', '')
    for vs_name in vs_info.items():
        for name in vs_name[1]:
            if tmp == name:
                ind = vs_name[1].index(name)
                code_list = vs_info['Код']
                test_vs_list = vs_info['Ссылка на описание тестового ВС']
                product_vs_list = vs_info['Ссылка на описание продуктивного ВС']
                if (test_vs_list[ind] != '') & (html.test_seg == "Yes"):
                    open_tab(html, test_vs_list[ind], 'TestS')
                if (product_vs_list[ind] != '') & (html.pr_seg == "Yes"):
                    open_tab(html, product_vs_list[ind], 'Product')
                return code_list[ind]
    return ""

def excel_work(html, vs_info):
    for vs_name in vs_info.items():
        for name in vs_name[1]:
            ind = vs_name[1].index(name)
            html.vs_name = name
            code_list = vs_info['Код']
            test_vs_list = vs_info['Ссылка на описание тестового ВС']
            product_vs_list = vs_info['Ссылка на описание продуктивного ВС']
            if ((str(test_vs_list[ind]) != 'nan') & (html.test_seg == "Yes")):
                make_dir(str(ROOT_DIR) + r"\downloads\\" + r"test_seg\\" + html.vs_name)
                open_tab(html, test_vs_list[ind], 'TestS')
            if ((str(product_vs_list[ind]) != 'nan') & (html.pr_seg == "Yes")):
                make_dir(str(ROOT_DIR) + r"\downloads\\" + r"pr_seg\\" + html.vs_name)
                open_tab(html, product_vs_list[ind], 'Product')
        if html.test_seg == "Yes":
            for dirs in os.listdir((str(ROOT_DIR) + r"\downloads\\" + r"test_seg\\")):
                for dir in os.listdir((str(ROOT_DIR) + r"\downloads\\" + r"test_seg\\" + dirs + "\\")):
                    print(dir)
                #comparing(str(ROOT_DIR) + r"\downloads\test_seg\\" + name)
        #if html.pr_seg == "Yes":
            #comparing(str(ROOT_DIR) + r"\downloads\pr_seg\\")
        break
    return ""

def download_by_xpath(html, xpath):
    try:
        html.browser.find_element_by_xpath(xpath).click()
        logger.info("Загрузка данных со страницы " + html.browser.current_url + ' ' + str(
            datetime.datetime.now()))
    except:
        logger.exception("Возникла ошибка при загрузке данных со страницы " + html.browser.current_url + ' ' + str(datetime.datetime.now()))
    downloads_done()
    time.sleep(1)

def replace_by_name(html, name):
    if str(name).endswith('test_seg.zip'):
        shutil.copyfile(name, str(ROOT_DIR) + r"\downloads\test_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.zip', '') + r'\\' + ntpath.basename(name), follow_symlinks=True)
        os.remove(name)
    if str(name).endswith('productive_seg.zip'):
        shutil.copyfile(name, str(ROOT_DIR) + r"\downloads\pr_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.zip', '') + r'\\' + ntpath.basename(name), follow_symlinks=True)
        os.remove(name)

def downloads_done():
    time.sleep(1)
    for i in os.listdir(str(ROOT_DIR) + r"\downloads"):
        if ".crdownload" in i:
            time.sleep(0.5)
            downloads_done()
        if ".tmp" in i:
            time.sleep(0.5)
            downloads_done()

def test():
    path = str(ROOT_DIR) + r"\information" + r"\Виды сведений.xlsx"
    vs_info = excel_connect(path)
    html = web('https://smev3.gosuslugi.ru/portal/news.jsp')
    check_vs(html, vs_info, path)

def newparser():
    test()
newparser()
