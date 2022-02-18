# -*- coding: UTF-8 -*-
import codecs
import configparser
import filecmp
import ast
import sys
import hashlib
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
        self.options.add_argument('headless')
        self.options.add_argument("--start-maximized")
        self.options.add_experimental_option("prefs", self.preferences)
        self.options.add_experimental_option('prefs', {
            "download.default_directory": str(ROOT_DIR) + r"\downloads",
            "download.directory_upgrade": True,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True
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


def edit_news(item):
    item = re.sub(r'href="#collapse....."', '', str(item))
    item = re.sub(r'<h3>', '<h2>', str(item))
    item = re.sub(r'<a class="collapsed" data-parent="#accordion" data-toggle="collapse" >',
                  '<span style="color:#FF0000" class="collapsed" data-parent="#accordion" data-toggle="collapse" >',
                  str(item))
    item = re.sub(r'</a></h3>', '</span> </h2>', str(item))
    return item


def get_themes(html):
    i = 2
    themes = []
    while True:
        xpath = '/html/body/div/div/div/div/div[1]/div/div[1]/div[2]/ul/li[i]/a'
        xpath = xpath.replace('[i]', '[' + str(i) + ']')
        ext = False
        if (len(html.browser.find_elements_by_xpath(xpath)) == 0) & (i == 2):
            ext = True
        try:
            if ext is False:
                wait = WebDriverWait(html.browser, 10, poll_frequency=1,
                                     ignored_exceptions=[NoSuchElementException, StaleElementReferenceException])
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                if (element.text == ''): break
                element.click()
            soup = BeautifulSoup(html.get_html_text(), "lxml")
            items = soup.find_all("div", {"class": "news-category"})
            check = False
            for item in items:
                for theme in themes:
                    if theme == item.get_text():
                        check = True
                        break
                if check is False:
                    themes.append(item.get_text())
                else:
                    check = False
            if ext is True:
               break
        except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
            break
        i += 1
        time.sleep(0.1)
    return themes


def check_file(file_name):
    for i in os.listdir(str(ROOT_DIR) + r"\downloads"):
        if file_name in i:
            return True
        continue
    return False

def rename_last_downloaded_zip(end_let):
    paths = sorted(Path(str(ROOT_DIR) + r"\downloads\\").iterdir(), key=lambda f: f.stat().st_mtime)
    filename = paths[len(paths)-1]
    try:
        if str(filename).endswith('.docx'):
            os.rename(filename, str(filename).replace('.docx', str(end_let).replace('.zip', '.docx')))
        if str(filename).endswith('.zip'):
            os.rename(filename, str(filename).replace('.zip', end_let))
    except FileExistsError:
        os.remove(filename)

def get_last_downloaded_file():
    paths = sorted(Path(str(ROOT_DIR) + r"\downloads\\").iterdir(), key=lambda f: f.stat().st_mtime)
    filename = paths[len(paths) - 1]
    return filename

def remain_2_last(path):
    paths = sorted(Path(path).iterdir(), key=lambda f: f.stat().st_mtime)
    for path in paths[:(len(paths)-2)]:
        os.remove(path)

def open_tab(html, link, mode):
    try:
        html.browser.execute_script("window.open();")
        main_window = html.browser.current_window_handle
        html.browser.switch_to.window(html.browser.window_handles[1])
        html.browser.get(link)
        time.sleep(0.1)
    except:
        main_window = ''
    if mode == 'Name':
        name = get_newsname(html)
    end_let = ''
    if mode == 'TestS':
        end_let = '_test_seg.zip'
    if mode == 'Product':
        end_let = '_productive_seg.zip'
    if end_let != '':
        time.sleep(0.1)
        """
        web_file = BeautifulSoup(html.browser.find_element_by_xpath('/html/body/div/div/div/div/div/div/div/div/table/tbody/tr[10]/td[1]/button').get_attribute('onclick'), 'lxml')
        try:
            found = re.search("files/(.+?)'", str(web_file.find('p'))).group(1)
        except AttributeError:
            found = ''
        print(found)
        """
        download_by_xpath(html, '//*[@id="infotable"]/table/tbody/tr[10]/td[1]/button')
        rename_last_downloaded_zip(end_let)
        last = get_last_downloaded_file()
        if mode == 'TestS':
            make = True
            for dir in os.listdir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name):
                if dir == ntpath.basename(str(last).replace('.zip', '')):
                    make = False
            if make is True:
                if ntpath.basename(str(last)).endswith('.zip') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name + "\\" + ntpath.basename(str(last).replace('.zip', '')))
                if ntpath.basename(str(last)).endswith('.docx') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.docx', '')))
        if mode == 'Product':
            make = True
            for dir in os.listdir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name):
                if dir == ntpath.basename(str(last).replace('.zip', '')):
                    make = False
            if make is True:
                if ntpath.basename(str(last)).endswith('.zip') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.zip', '')))
                if ntpath.basename(str(last)).endswith('.docx') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.docx', '')))
        replace_by_name(html, last)
        download_by_xpath(html, '//*[@id="infotable"]/table/tbody/tr[11]/td[1]/button')
        rename_last_downloaded_zip(end_let)
        last = get_last_downloaded_file()
        if mode == 'TestS':
            make = True
            for dir in os.listdir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name):
                if dir == ntpath.basename(str(last).replace('.zip', '')):
                    make = False
            if make is True:
                if ntpath.basename(str(last)).endswith('.zip') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.zip', '')))
                if ntpath.basename(str(last)).endswith('.docx') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.docx', '')))
        if mode == 'Product':
            make = True
            for dir in os.listdir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name):
                if dir == ntpath.basename(str(last).replace('.zip', '')):
                    make = False
            if make is True:
                if ntpath.basename(str(last)).endswith('.zip') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.zip', '')))
                if ntpath.basename(str(last)).endswith('.docx') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.docx', '')))
        replace_by_name(html, last)
        download_by_xpath(html, '//*[@id="infotable"]/table/tbody/tr[12]/td[1]/button')
        rename_last_downloaded_zip(end_let)
        last = get_last_downloaded_file()
        if mode == 'TestS':
            make = True
            for dir in os.listdir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name):
                if dir == ntpath.basename(str(last).replace('.zip', '')):
                    make = False
            if make is True:
                if ntpath.basename(str(last)).endswith('.zip') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.zip', '')))
                if ntpath.basename(str(last)).endswith('.docx') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\test_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.docx', '')))
        if mode == 'Product':
            make = True
            for dir in os.listdir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name):
                if dir == ntpath.basename(str(last).replace('.zip', '')):
                    make = False
            if make is True:
                if ntpath.basename(str(last)).endswith('.zip') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.zip', '')))
                if ntpath.basename(str(last)).endswith('.docx') is True:
                    make_dir(str(ROOT_DIR) + r"\downloads\\" + r"\pr_seg\\" + html.vs_name + "\\" + ntpath.basename(
                        str(last).replace('.docx', '')))
        replace_by_name(html, last)
    html.browser.close()
    html.browser.switch_to.window(main_window)
    time.sleep(0.1)
    if mode == 'Name':
        return name

def make_dir(name):
    try:
        os.mkdir(name)
    except:
        pass

def get_newsname(html):
    try:
        element = WebDriverWait(html.browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="infotable"]/table/tbody/tr[1]'))
        )
        name = element.text
        name = re.sub(r'Наименование ', '', str(name))
        name = '<h3> ' + 'Новость содержит информацию о следующем ВС - ' + name + ' </h3>'
    except TimeoutException:
        name = ''
    return name

def make_html_text(text_list):
    name = ['<h2>']
    for text in text_list:
        name.append(text + ' ' + '\n')
    name.append('</h2>')
    return name

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
                    comparing_for_vs((str(ROOT_DIR) + r"\downloads\\" + r"test_seg\\" + dirs + "\\") + dir, 'test')
            if test_info_comp != []:
                for mail in html.dest_mails:
                    send_email(make_html_text(test_info_comp), mail, "Тестовый сегмент_" + html.mail_subject, "comparing", html.send_addr, html.send_password)
                    #comparing(str(ROOT_DIR) + r"\downloads\test_seg\\" + name)
        if html.pr_seg == "Yes":
            for dirs in os.listdir((str(ROOT_DIR) + r"\downloads\\" + r"pr_seg\\")):
                for dir in os.listdir((str(ROOT_DIR) + r"\downloads\\" + r"pr_seg\\" + dirs + "\\")):
                    comparing_for_vs((str(ROOT_DIR) + r"\downloads\\" + r"pr_seg\\" + dirs + "\\") + dir, 'pr')
            if pr_info_comp != []:
                for mail in html.dest_mails:
                    send_email(make_html_text(pr_info_comp), mail, "Продуктивный сегмент_" + html.mail_subject, "comparing", html.send_addr, html.send_password)
                #comparing(str(ROOT_DIR) + r"\downloads\pr_seg\\")
        break
    return ""

def comparing_for_vs(path, type):
    files_in_dir = []
    last = ''
    for root, dirs, files in os.walk(path):
        for file in files:
            for cmp in files:
                if (cmp != file) & (cmp != last):
                    filecmp.clear_cache()
                    ch = (filecmp.cmp(path + '\\' + file, path + '\\' + cmp, shallow=False))
                    #filecompare(path + '\\' + file, path + '\\' + cmp)
                    if ch is False:
                        z1_u = zip_unicode.ZipHandler(path + '\\' + file)
                        z2_u = zip_unicode.ZipHandler(path + '\\' + cmp)
                        z1_u.extract_all()
                        z2_u.extract_all()
                        dcmp = filecmp.dircmp(path + '\\' + file.replace('.zip', ''),
                                              path + '\\' + cmp.replace('.zip', ''))
                        for fd, subfds, fns in os.walk(path + '\\' +
                                                   file.replace('.zip', '')):
                            for fn in fns:
                                 for fd2, subfds2, fns2 in os.walk(path + '\\' + cmp.replace('.zip', '')):
                                    for fn2 in fns2:
                                        if similar(fn, fn2) >= 0.95:
                                            if filecmp.cmp(os.path.join(fd, fn),
                                                       os.path.join(fd2, fn2),
                                                       shallow=False) is False:
                                                if type == 'test':
                                                    test_info_comp.append("Файл " + str(fn) + " в " + str(file) + " изменен.")
                                                if type == 'pr':
                                                    pr_info_comp.append(
                                                        "Файл " + str(fn) + " в " + str(file) + " изменен.")
                        shutil.rmtree(path + '\\' + file.replace('.zip', ''))
                        shutil.rmtree(path + '\\' + cmp.replace('.zip', ''))
            last = file

def download_by_xpath(html, xpath):
    try:
        html.browser.find_element_by_xpath(xpath).click()
        logger.info("Загрузка данных со страницы " + html.browser.current_url + ' ' + str(
            datetime.datetime.now()))
    except:
        logger.exception("Возникла ошибка при загрузке данных со страницы " + html.browser.current_url + ' ' + str(datetime.datetime.now()))
    downloads_done()
    time.sleep(1)

def check_news(name):
    if os.stat(str(ROOT_DIR) + r"\information" + r"\news.txt").st_size == 0:
        f = open(str(ROOT_DIR) + r"\information" + r"\news.txt", 'w')
        f.write(name + '\n')
        return True
    f = open(str(ROOT_DIR) + r"\information" + r"\news.txt", 'r')
    for line in f:
        if str(line).replace('\n', '') == str(name).replace('\n', ''):
            return False
    f = open(str(ROOT_DIR) + r"\information" + r"\news.txt", 'a')
    f.write(name + '\n')
    f.close()
    return True

def get_xml(html, link):
    try:
        html.browser.execute_script("window.open();")
        main_window = html.browser.current_window_handle
        html.browser.switch_to.window(html.browser.window_handles[1])
        html.browser.get(link)
        time.sleep(0.1)
    except:
        main_window = ''
    xml = html.browser.find_element_by_xpath('/html/body/pre')
    make_file(xml.text)
    html.browser.close()
    html.browser.switch_to.window(main_window)
    time.sleep(0.1)

def get_files(input):
    for fd, subfds, fns in os.walk(input):
       for fn in fns:
            yield os.path.join(fd, fn)

def compare_files_with_text(text):
    for filename in os.listdir(str(ROOT_DIR) + r"\downloads"):
        if filename.endswith(".xsd") and open(str(ROOT_DIR) + r"\downloads\\" + filename, "r", encoding='utf-8').read() == text:
            return False
    return True

def make_file(text):
    if compare_files_with_text(text):
        path = str(ROOT_DIR) + r"\downloads\\"
        uniq_filename = path + 'xml_code_' + str(datetime.datetime.now().date()) + '-' +str(datetime.datetime.now().time()).replace(':','-') + '.xsd'
        text_file = open(uniq_filename, "w", encoding='utf-8')
        text_file.write(text)

def check_jkh(name):
    if os.stat(str(ROOT_DIR) + r"\information" + r"\jkh_name.txt").st_size == 0:
        f = open(str(ROOT_DIR) + r"\information" + r"\jkh_name.txt", 'w')
        f.write(name + '\n')
        return True
    f = open(str(ROOT_DIR) + r"\information" + r"\jkh_name.txt", 'r')
    for line in f:
        if str(line).replace('\n', '') == str(name).replace('\n', ''): return False
    f = open(str(ROOT_DIR) + r"\information" + r"\jkh_name.txt", 'a')
    f.write(name + '\n')
    return True

def renamed(dirpath, names, encoding):
    new_names = [old.encode('cp437').decode(encoding) for old in names]
    for old, new in zip(names, new_names):
        os.rename(os.path.join(dirpath, old), os.path.join(dirpath, new))
    return new_names

def change_name(name):
    import chardet
    try:
        name = name.encode('cp437')
    except UnicodeEncodeError:
        name = name.encode('utf8')
    encoding = chardet.detect(name)['encoding']
    name = name.decode(encoding)
    return name

def replace_by_name(html, name):
    if str(name).endswith('test_seg.zip'):
        shutil.copyfile(name, str(ROOT_DIR) + r"\downloads\test_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.zip', '') + r'\\' + str(ntpath.basename(name)).replace('.zip', str(datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")) + '.zip'), follow_symlinks=True)
        os.remove(name)
        remain_2_last(str(ROOT_DIR) + r"\downloads\test_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.zip', ''))
    if str(name).endswith('productive_seg.zip'):
        shutil.copyfile(name, str(ROOT_DIR) + r"\downloads\pr_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.zip', '') + r'\\' + str(ntpath.basename(name)).replace('.zip', str(datetime.datetime.now().strftime("%d-%m-%Y%H-%M-%S")) + '.zip'), follow_symlinks=True)
        os.remove(name)
        remain_2_last(str(ROOT_DIR) + r"\downloads\pr_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.zip', ''))
    if str(name).endswith('test_seg.docx'):
        shutil.copyfile(name, str(ROOT_DIR) + r"\downloads\test_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.docx', '') + r'\\' + str(ntpath.basename(name)).replace('.docx', str(datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")) + '.docx'), follow_symlinks=True)
        os.remove(name)
        remain_2_last(str(ROOT_DIR) + r"\downloads\test_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.docx', ''))
    if str(name).endswith('productive_seg.docx'):
        shutil.copyfile(name, str(ROOT_DIR) + r"\downloads\pr_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.docx', '') + r'\\' + str(ntpath.basename(name)).replace('.docx', str(datetime.datetime.now().strftime("%d-%m-%Y%H-%M-%S")) + '.docx'), follow_symlinks=True)
        os.remove(name)
        remain_2_last(str(ROOT_DIR) + r"\downloads\pr_seg\\" + html.vs_name + r'\\' + str(ntpath.basename(name)).replace('.docx', ''))


def comparing(path):
    files_in_dir = []
    for root, dirs, files in os.walk(path):
        if str(root).find('pr_seg') >= 0: continue
        if str(root).find('test_seg') >= 0: continue
        for filename in files:
            if check_temp(filename):
                for dir_file in files_in_dir:
                    if similar(filename, dir_file) >= 0.9:
                        if filename.endswith('.zip') & dir_file.endswith('.zip'):
                            if filecmp.cmp(path + filename,
                                           path + dir_file, shallow=False) is False:
                                if dir_file.startswith('AF.2.65d') or filename.startswith('AF.2.65d'):
                                    info_comp.append('В архиве ' + filename + ' замечены изменеия.')
                                else:
                                    z1_u = zip_unicode.ZipHandler(path + filename)
                                    z2_u = zip_unicode.ZipHandler(path + dir_file)
                                    z1_u.extract_all()
                                    z2_u.extract_all()
                                    dcmp = filecmp.dircmp(path + filename.replace('.zip', ''),
                                                          path + dir_file.replace('.zip', ''))
                                    if dcmp.left_only != []:
                                        for lefts in dcmp.left_only:
                                            if path == str(ROOT_DIR) + r"\downloads\pr_seg\\":
                                                pr_info_comp.append(
                                                    'В архиве ' + filename + ' присутсвует файл ' + lefts + ', который отсутсвует в ' + dir_file + ' ' + '\n')
                                            if path == str(ROOT_DIR) + r"\downloads\test_seg\\":
                                                test_info_comp.append(
                                                    'В архиве ' + filename + ' присутсвует файл ' + lefts + ', который отсутсвует в ' + dir_file + ' ' + '\n')
                                            else: info_comp.append('В архиве ' + filename + ' присутсвует файл ' + lefts + ', который отсутсвует в ' + dir_file + ' ' + '\n')
                                    if dcmp.right_only != []:
                                        for rights in dcmp.right_only:
                                            info_comp.append('В архиве ' + dir_file + 'присутсвует файл ' + rights + ', который отсутсвует в ' + filename + ' ' + '\n')
                                    for fd, subfds, fns in os.walk(path +
                                                                           filename.replace('.zip', '')):
                                        for fn in fns:
                                            for fd2, subfds2, fns2 in os.walk(path + dir_file.replace('.zip', '')):
                                                for fn2 in fns2:
                                                    if similar(fn, fn2) >= 0.95:
                                                        if filecmp.cmp(os.path.join(fd,fn),
                                                                       os.path.join(fd2, fn2),
                                                                       shallow=False) is False:
                                                            if fn.endswith('.docx') & fn2.endswith('.docx') & (fn.endswith('_Comparison.docx') is False) & (fn2.endswith('_Comparison.docx') is False) & (fn.startswith('~$') is False) & (fn2.startswith('~$') is False):
                                                                compare_name = ntpath.basename(compare_docs(fn, fn2, os.path.join(fd,fn),os.path.join(fd2, fn2)))
                                                                info_comp.append(
                                                                    'В архиве ' + dir_file + ' обнаружены изменения файла ' + fn + ' изменения записаны в прикрепленный файл ' + compare_name + '\n')
                                shutil.rmtree(path + filename.replace('.zip', ''))
                                shutil.rmtree(path + dir_file.replace('.zip', ''))
                            else:
                                os.remove(path + filename)
                        else:
                            if filename.endswith('.docx') & dir_file.endswith('.docx') & (filename.endswith('_Comparison.docx') is False) & (dir_file.endswith('_Comparison.docx') is False):
                                try:
                                    info_comp.append('В файле ' + filename + ' присутствуют изменения, которые записаны в файл ' + ntpath.basename(compare_docs(change_name(filename), change_name(dir_file), str(ROOT_DIR) + r"\downloads\\" , str(ROOT_DIR) + r"\downloads\\")) + '\n')
                                except:
                                    logger.exception("Возникла ошибка при работе с файлами " + filename + ' и ' + dir_file + ' ' + str(datetime.datetime.now()))
                files_in_dir.append(filename)

def get_content(html):
    news = []
    if html.link == 'https://smev3.gosuslugi.ru/portal':
        soup = BeautifulSoup(html.get_html_text(), "lxml")
        items = soup.find_all("div", {"class": "info-section gray-container"})
        for item in items:
            item_tmp = item
            color_check1 = item_tmp.find("span", style='color:#FF0000')
            color_check2 = item_tmp.find("span", style='color:rgb(255, 0, 0)')
            if (color_check1 != None or color_check2 != None):
                downloads = item.find_all("span", {"class": "is__filename"})
                for download in downloads:
                    download = download.find('a').get('href')
                    regex = r'[^/\\&\?]+\.\w{3,4}(?=([\?&].*$|$))'
                    file_name = re.search(regex, download).group(0)
                    if check_file(file_name) == False:
                        try:
                            html.browser.find_element_by_xpath('//a[@href="' + download + '"]').click()
                        except BaseException:
                            html.browser.back()
                            continue
                        downloads_done()
                tmps = item.find_all('h2') + item.find_all('h3') + item.find_all('p') + item.find_all('h4')
                news_body = ''
                for i in range(0, len(tmps)):
                    news_body = (str(news_body) + ''.join(map(str, tmps[i]))).replace('\n', '')
                if check_news(news_body) is True:
                    news.append(tmps)
    if html.link == 'https://smev3.gosuslugi.ru/portal/news.jsp':
        try:
            path = str(ROOT_DIR) + r"\information" + r"\Виды сведений.xlsx"
            vs_info = excel_connect(path)
        except:
            path = ''
            vs_info = ''
        i = 2
        while True:
            xpath = '/html/body/div/div/div/div/div[1]/div/div[1]/div[2]/ul/li[i]/a'
            xpath = xpath.replace('[i]', '[' + str(i) + ']')
            ext = False
            if (len(html.browser.find_elements_by_xpath(xpath)) == 0) & (i == 2):
                ext = True
            try:
                if ext is False:
                    wait = WebDriverWait(html.browser, 10, poll_frequency=1, ignored_exceptions=[NoSuchElementException, StaleElementReferenceException])
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    if (element.text == ''): break
                    element.click()
                soup = BeautifulSoup(html.get_html_text(), "lxml")
                items = soup.find_all("div", {"class": "panel panel-news"})
                for item in items:
                    links_with_text = []
                    for a in item.find_all('a', href=True):
                        if a.text:
                            if str(a['href']).startswith('https://smev3.gosuslugi.ru') == True:
                                if(similar(a.text, "ссылке") >= 0.9):
                                    links_with_text.append(a['href'])
                    if links_with_text != []:
                        name = open_tab(html, links_with_text[0], 'Name')
                    else:
                        name = ''
                    vs_add = check_vs(html, vs_info, name)
                    if vs_add != '':
                        name = name + "Данные найдены в таблице видов сведений со следующим кодом - " + vs_add
                    else:
                        name = name + "Данные не найдены в таблице видов сведений"
                    news_body = ''
                    if html.filter == True:
                        if item.get_text().find(html.filter_word) != -1:
                            main_news = edit_news(item)
                            news_body = (str(news_body) + ''.join(map(str, main_news))).replace('\n', '')
                            if check_news(news_body) is True:
                                news.append(name + ' ' + main_news)
                    else:
                        main_news = edit_news(item)
                        news_body = (str(news_body) + ''.join(map(str, main_news))).replace('\n', '')
                        if check_news(news_body) is True:
                            news.append(name + ' ' + main_news)
                if ext is True: break
            except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                break
            i += 1
            time.sleep(0.1)
        excel_work(html, vs_info)
        #for file in os.listdir(str(ROOT_DIR) + r"\downloads\\"):
         #   replace_by_name(str(ROOT_DIR) + r"\downloads\\" + file)
    if html.link == 'https://dom.gosuslugi.ru/#!/regulations':
        i = 1
        j = 2
        tmp = "Регламент и форматы информационного взаимодействия внешних информационных систем с ГИС ЖКХ (текущие"
        while True:
            xpath = '// *[ @ id = "rubric_9}"] / div[1] / div[i] / div / div / div / div[1] / div[1] / a'
            xpath = xpath.replace('[i]', '[' + str(i) + ']')
            try:
                wait = WebDriverWait(html.browser, 10, poll_frequency=2,ignored_exceptions=[NoSuchElementException, StaleElementReferenceException])
                element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                if element.text.startswith(tmp) is True:
                    if check_jkh(element.text) is True:
                        element.click()
                        downloads_done()
            except:
                break
            if i == 5:
                i = 0
                xpath = '//*[@id="rubric_9}"]/div[2]/div[1]/div/ul[2]/li[i]/a'
                xpath = xpath.replace('[i]', '[' + str(j) + ']')
                try:
                    element = html.browser.find_element_by_xpath(xpath)
                    element.click()
                    time.sleep(0.1)
                except:
                    break
            i += 1
            time.sleep(0.1)
    if html.link == 'https://fssp.gov.ru/mvv_fssp/':
        soup = BeautifulSoup(html.get_html_text(), "lxml")
        items = soup.find("div", {"class": "b-responsive-table"})
        tbody = items.find('tbody')
        trs = tbody.find_all('tr')
        td_check = False
        for tr in trs:
            tds = tr.find_all('td')
            for td in tds:
                test = 'Запросы должностных лиц ФССП России и ответы на них'
                if similar(test, td.text) >= 0.95:
                    td_check = True
                else:
                    if td_check is True:
                        for a in td.find_all('a', href=True):
                            test1 = 'загрузить pdf'
                            test2 = 'загрузить xsd'
                            if similar(a.text,test1) >= 0.9:
                                regex = r'[^/\\&\?]+\.\w{3,4}(?=([\?&].*$|$))'
                                file_name = re.search(regex, a['href']).group(0)
                                if check_file(file_name) is False:
                                    open_tab(html, 'https://fssp.gov.ru/' + a['href'], '')
                            if similar(a.text, test2) >= 0.9:
                                get_xml(html, 'https://fssp.gov.ru/' + a['href'])
                        downloads_done()
                        td_check = False
    if html.link == 'https://pfr.gov.ru/info/af/':
        soup = BeautifulSoup(html.get_html_text(), "lxml")
        items = soup.find("div", {"id": "accordion"})
        tmp = ('Альбомформатов 2.65д')
        for a in items.find_all('a', href=True):
            if re.sub("^\s+|\n|\r|\s+$", '', a.text).replace('	', '').startswith(tmp):
                open_tab(html, 'https://pfr.gov.ru' + a['href'], '')
                downloads_done()
    return news


def downloads_done():
    time.sleep(1)
    for i in os.listdir(str(ROOT_DIR) + r"\downloads"):
        if ".crdownload" in i:
            time.sleep(0.5)
            downloads_done()
        if ".tmp" in i:
            time.sleep(0.5)
            downloads_done()


def send_email(news, toaddr, subject, msg_type, sender, passw):
    email_str = sender
    password = passw

    server = smtplib.SMTP('smtp.yandex.ru', 587)
    server.set_debuglevel(False)
    server.ehlo()
    server.starttls()
    server.login(email_str, password)

    from_addr = email_str
    msg_body = '<html> <body>'
    for i in range(0, len(news)):
        msg_body = str(msg_body) + ''.join(map(str, news[i]))
    msg_body = str(msg_body) + '</body></html>'
    msg = MIMEMultipart('alternative')
    part = MIMEText(str(msg_body), 'html', 'utf-8')
    msg.attach(part)
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = toaddr
    if msg_type == "comparing":
        if files_to_send != []:
            last = ''
            for file_n in files_to_send:
                if file_n != last:
                    attach = MIMEApplication(open(file_n, 'rb').read())
                    attach.add_header('Content-Disposition', 'attachment', filename = ntpath.basename(file_n))
                    msg.attach(attach)
                last = file_n
    try:
        server.sendmail(email_str, toaddr, msg.as_string())
        logger.info("Сообщение успешно отправлено")
    except (SMTPDataError, smtplib.SMTPSenderRefused):
        logger.exception("Сообщение не доставлено, проверьте работоспособность исходящего адреса, размер вложения слишком велик.")
    server.quit()

def end_delete():
    if files_to_send != []:
        for file_n in files_to_send:
            try:
                os.remove(file_n)
            except: return ""

def check_theme(theme_name, themes):
    for theme in themes:
        if theme_name == theme:
            return True
    return False

def check_temp(name):
    if str(name).startswith('~$'):
        return False
    else:
        return True


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def compare_docs(doc1, doc2, path1 , path2):
    path = str(ROOT_DIR) + r"\downloads\\"
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = True
    path1 = path1.replace(str(doc1), '') + r"\\"
    path2 = path2.replace(str(doc2), '') + r"\\"
    word.CompareDocuments(word.Documents.Open(path1 + str(doc1)),
                                 word.Documents.Open(path2 + str(doc2)))
    name = str(doc1).replace('.docx', '') + "_Comparison.docx"
    word.ActiveDocument.SaveAs(FileName=path + name)
    word.Quit()
    file_name = path + name
    file_name = file_name.replace(r'\\\\', r'\\')
    files_to_send.append(file_name)
    word.Quit()
    return file_name

def init_delete():
    for root, dirs, files in os.walk(str(ROOT_DIR) + r"\downloads"):
        for filename in files:
            if filename.endswith("_Comparison.docx"):
                os.remove(str(ROOT_DIR) + r"\downloads\\" + filename)
            else:
                if filename.endswith("(1).docx"):
                    os.remove(str(ROOT_DIR) + r"\downloads\\" + filename)

def test():
    path = str(ROOT_DIR) + r"\information" + r"\Виды сведений.xlsx"
    vs_info = excel_connect(path)
    html = web('https://smev3.gosuslugi.ru/portal/news.jsp')
    check_vs(html, vs_info, path)

def newparser():
    logger.info("Программа запущена. " + str(datetime.datetime.now()))
    init_delete()
    config = configparser.ConfigParser()
    try:
        if os.stat(str(ROOT_DIR) + r"\\config.ini").st_size == 0:
            shutil.copy(str(ROOT_DIR) + r"\\config_reserve.ini", str(ROOT_DIR) + r"\\config.ini")
        config.read(str(ROOT_DIR) + r"\\config.ini", 'UTF-8')
        my_list = ast.literal_eval(config.get("links", "link"))
        dest_mails = ast.literal_eval(config.get("e_mail", "destination_addr"))
        mail_subject = config["e_mail"]["msg_subject"]
        send_addr = config["e_mail"]["send_addr"]
        send_password = config["e_mail"]["send_password"]
        test_seg = config["excel"]["test_ct"]
        pr_seg = config["excel"]["product_ct"]
    except:
        my_list = []
        dest_mails = []
        mail_subject = ""
        send_addr = ''
        send_password = ''
        test_seg = ""
        pr_seg = ""
        logger.exception("Не удалось считать данный из конфигурационного файла. " + str(datetime.datetime.now()))
    for link in my_list:
        smev3_news = []
        send_news = False
        if link == 'https://smev3.gosuslugi.ru/portal/news.jsp' or link == 'https://smev3.gosuslugi.ru/portal':
            send_news = True
        html = web(link)
        html.mail_subject = mail_subject
        html.dest_mails = dest_mails
        html.send_addr = send_addr
        html.send_password = send_password
        if link == 'https://smev3.gosuslugi.ru/portal/news.jsp':
            themes = get_themes(html)
            to_cfg = '' 
            for th in themes:
                to_cfg += (th + ', ')
            config.set('filter', 'themes', to_cfg)
            with codecs.open('config.ini', 'w', 'UTF-8') as configfile:
                config.write(configfile)
            if config["filter"]["filter_on"] == '"Yes"':
                html.filter = True
                html.filter_word = config["filter"]["theme"]
                if check_theme(html.filter_word, themes) == False:
                    logger.info("Тема фильтрации не обнаружена, отбор новостей осуществить не удалось")
                    del html
                    continue
        else:
            html.filter = False
        html.pr_seg = pr_seg
        html.test_seg = test_seg
        smev3_news += get_content(html)
        if send_news is True:
            if smev3_news != []:
                for mail in dest_mails:
                    send_email(smev3_news, mail, mail_subject, "news", send_addr, send_password)
            else:
                logger.info("На странице " + link + " новых новостей не обнаружено. " + str(datetime.datetime.now()))
        del html
        time.sleep(0.1)
    comparing(str(ROOT_DIR) + r"\downloads\\")
    if info_comp != []:
        jkh_cmp = []
        other_cmp = []
        last = ''
        for mail in dest_mails:
            for cmp in info_comp:
                if cmp != last:
                    if str(cmp).find("ГИС") >= 0:
                        jkh_cmp.append(cmp)
                    else:
                        other_cmp.append(cmp)
                last = cmp
            if jkh_cmp != []:
                send_email(make_html_text(jkh_cmp), mail, mail_subject, "jkh", send_addr, send_password)
            if other_cmp != []:
                send_email(make_html_text(other_cmp), mail, mail_subject, "comparing", send_addr, send_password)
    else:
        logger.info("Новых версий файлов не обнаружено." + str(datetime.datetime.now()))
    end_delete()
    logger.info("Программа выполнена. " + str(datetime.datetime.now()) + '\n')
    #test()
newparser()
