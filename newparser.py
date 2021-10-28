# -*- coding: cp1251 -*-
import os
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
import ntpath
import datetime
"""
from selenium.webdriver import ActionChains

actions = ActionChains(driver)
actions.move_to_element(element).click().perform()
"""

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
chromedriver = ROOT_DIR + r"\chromedriver"
files_to_send = []

class web:

    def __init__(self, link):
        self.options = webdriver.ChromeOptions()
        self.preferences = ''
        self.set_preferencies()
        self.set_options()
        self.browser = webdriver.Chrome(executable_path=chromedriver, options=self.options)
        self.msg = MIMEMultipart()
        self.link = link
        self.connection(link)
        self.filter = False
        self.filter_word = ''

    def set_options(self):
        # self.options.add_argument('headless')
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
        print('Отключение от сайта ' + self.link)
        try:
            self.browser.close()
            print('Отключение прошло успешною')
        except WebDriverException:
            print('Произошла ошибка при отключении: ' + traceback.format_exc())
            input("Для продолжения нажмите Enter")
    def get_html_text(self):
        return self.browser.page_source

    def connection(self, link):
        print("Подключение к сайту " + link + "...")
        try:
            self.browser.get(link)
            time.sleep(1)
            print("Подключение прошло успешно")
        except WebDriverException:
            print("Возникла ошибка при подключении:")
            print(traceback.format_exc())
            input("Для продолжения нажмите Enter")
        return self.browser.page_source

    def __del__(self):
        self.close_connection()


def excel_connect(path):
    pandas.set_option('display.max_rows', None)
    pandas.set_option('display.max_columns', None)
    pandas.set_option('display.max_colwidth', None)

    vs_info = pandas.read_excel(path, index_col=None, na_values=['NA'], usecols="F,E")
    vs_info_dict = vs_info.set_index('Код').to_dict()
    vs_info_dict = vs_info_dict['Название Сервиса(ВС) в СМЭВ']
    return vs_info_dict


def edit_news(item):
    item = re.sub(r'href="#collapse....."', '', str(item))
    item = re.sub(r'<h3>', '<h2>', str(item))
    item = re.sub(r'<a class="collapsed" data-parent="#accordion" data-toggle="collapse" >',
                  '<span style="color:#FF0000" class="collapsed" data-parent="#accordion" data-toggle="collapse" >',
                  str(item))
    item = re.sub(r'</a></h2>', '</span></h2>', str(item))
    return item


def get_themes(html):
    i = 2
    themes = []
    while True:
        xpath = '/html/body/div/div/div/div/div[1]/div/div[1]/div[2]/ul/li[i]/a'
        xpath = xpath.replace('[i]', '[' + str(i) + ']')
        try:
            html.browser.find_element_by_xpath(xpath).click()
            soup = BeautifulSoup(html.get_html_text(), "lxml")
            items = soup.find_all("div", {"class": "news-category"})
            check = False
            for item in items:
                for theme in themes:
                    if theme == item.get_text():
                        check = True
                        break
                if check == False:
                    themes.append(item.get_text())
                else:
                    check = False
        except (NoSuchElementException, ElementClickInterceptedException):
            return themes
        i += 1
        time.sleep(0.1)


def check_file(file_name):
    for i in os.listdir(str(ROOT_DIR) + r"\downloads"):
        if file_name in i:
            return True
        continue
    return False


def open_tab(html, link, need_name):
    try:
        html.browser.execute_script("window.open();")
        main_window = html.browser.current_window_handle
        html.browser.switch_to.window(html.browser.window_handles[1])
        html.browser.get(link)
        time.sleep(0.1)
    except:
        print(traceback.format_exc())
    if need_name is True:
        name = get_newsname(html)
    html.browser.close()
    html.browser.switch_to.window(main_window)
    time.sleep(0.1)
    if need_name is True:
        return name


def get_newsname(html):
    wait = WebDriverWait(html.browser, 10, poll_frequency=1, ignored_exceptions=[NoSuchElementException, StaleElementReferenceException])
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="infotable"]/table/tbody/tr[1]')))
    name = element.text
    name = re.sub(r'Наименование ', '', str(name))
    name = '<h3> ' + 'Новость содержит информацию о следующем ВС - ' + name + ' </h3>'
    return name


def check_vs(vs_info, name):
    for code, vs_name in vs_info.items():
        if name == vs_name:
            return code
        return ""

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
        print(traceback.format_exc())
    xml = html.browser.find_element_by_xpath('/html/body/pre')
    make_file(xml.text)
    html.browser.close()
    html.browser.switch_to.window(main_window)
    time.sleep(0.1)

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
        files_in_dir = []
        for root, dirs, files in os.walk(str(ROOT_DIR) + r"\downloads"):
            for filename in files:
                if check_temp(filename):
                    for dir_file in files_in_dir:
                        if similar(filename, dir_file) >= 0.7:
                            compare_docs(filename, dir_file)
                    files_in_dir.append(filename)
        return news
    if html.link == 'https://smev3.gosuslugi.ru/portal/news.jsp':
        path = str(ROOT_DIR) + r"\information" + r"\Виды сведений.xlsx"
        vs_info = excel_connect(path)
        i = 2
        while True:
            xpath = '/html/body/div/div/div/div/div[1]/div/div[1]/div[2]/ul/li[i]/a'
            xpath = xpath.replace('[i]', '[' + str(i) + ']')
            try:
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
                                links_with_text.append(a['href'])
                    if links_with_text != []:
                        name = open_tab(html, links_with_text[0], True)
                    else:
                        name = ''
                    vs_add = check_vs(vs_info, name)
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
            except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                print(traceback.format_exc())
                break
            i += 1
            time.sleep(0.1)
    if html.link == 'https://dom.gosuslugi.ru/#!/regulations':
        i = 1
        j = 2
        tmp = "Регламент и форматы информационного взаимодействия внешних информационных систем с ГИС ЖКХ (текущие форматы"
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
            except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                break
            if i == 5:
                i = 0
                xpath = '//*[@id="rubric_9}"]/div[2]/div[1]/div/ul[2]/li[i]/a'
                xpath = xpath.replace('[i]', '[' + str(j) + ']')
                try:
                    element = html.browser.find_element_by_xpath(xpath)
                    element.click()
                    time.sleep(0.1)
                except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                    break
            i += 1
            time.sleep(0.1)
        return news
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
                                    open_tab(html, 'https://fssp.gov.ru/' + a['href'], False)
                            if similar(a.text, test2) >= 0.9:
                                get_xml(html, 'https://fssp.gov.ru/' + a['href'])
                        downloads_done()
                        td_check = False
    if html.link == 'https://pfr.gov.ru/info/af/':
        soup = BeautifulSoup(html.get_html_text(), "lxml")
        items = soup.find("div", {"id": "accordion"})
        tmp = 'Альбомформатов 2.64д'
        for a in items.find_all('a', href=True):
            if re.sub("^\s+|\n|\r|\s+$", '', a.text).replace('	', '').startswith(tmp):
                open_tab(html, 'https://pfr.gov.ru' + a['href'], False)
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


def send_email(news, toaddr):
    email_str = 'tesmail.test@yandex.ru'
    password = 'Pinkuin5'

    server = smtplib.SMTP('smtp.yandex.ru', 587)
    server.set_debuglevel(False)
    server.ehlo()
    server.starttls()
    server.login(email_str, password)

    from_addr = 'tesmail.test@yandex.ru'
    msg_body = '<html> <body>'
    for i in range(0, len(news)):
        msg_body = str(msg_body) + ''.join(map(str, news[i]))
    msg_body = str(msg_body) + '</body></html>'

    msg = MIMEMultipart('alternative')
    part = MIMEText(str(msg_body), 'html', 'utf-8')
    msg.attach(part)
    msg['Subject'] = 'Это письмо от Песоцкого Дмитрия'
    msg['From'] = from_addr
    msg['To'] = toaddr

    if files_to_send != []:
        for file_n in files_to_send:
            attach = MIMEApplication(open(file_n, 'rb').read())
            attach.add_header('Content-Disposition', 'attachment', filename = ntpath.basename(file_n))
            msg.attach(attach)

    server.sendmail(email_str, toaddr, msg.as_string())
    print("Сообщение успешно отправлено")
    server.quit()


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

def compare_docs(doc1, doc2):
    path = str(ROOT_DIR) + "\downloads\\"
    word = win32com.client.Dispatch("Word.application")
    word.CompareDocuments(word.Documents.Open(path + str(doc1)), word.Documents.Open(path + str(doc2)))
    doc1 = str(doc1).replace(".docx", '')
    doc2 = str(doc2).replace(".docx", '')
    word.ActiveDocument.ActiveWindow.View.Type = 3
    file_name = path + doc1 + "_Comparison.docx"
    word.ActiveDocument.SaveAs(FileName=file_name, FileFormat = 16)
    files_to_send.append(file_name)
    word.Quit()
    return file_name

def newparser():
    html = web('https://smev3.gosuslugi.ru/portal')
    smev3_news = get_content(html)
    del html
    time.sleep(0.1)
    html_1 = web('https://smev3.gosuslugi.ru/portal/news.jsp')
    print("На данный момент доступны следующие фильтры для сайта https://smev3.gosuslugi.ru/portal/news.jsp: ")
    themes = get_themes(html_1)
    for theme in themes:
        print('-- ' + str(theme))
    answer = input("Желаете ли Вы воспользоваться фильтром? (Да/Нет): ")
    if answer == 'Да':
        html_1.filter = True
        html_1.filter_word = input("Скопируйте тему фильтрации из списка выше сюда: ")
        while check_theme(html_1.filter_word, themes) == False:
            print("Тема фильтрации не соответсвует списку")
            html_1.filter_word = input("Скопируйте тему фильтрации из списка выше сюда: ")
    else:
        html_1.filter = False
    print("Производится отбор новостей по заданным параметрам...")
    next_smev3_news = get_content(html_1)
    del html_1
    print("Новости отобраны.")
    html_2 = web('https://dom.gosuslugi.ru/#!/regulations')
    get_content(html_2)
    del html_2
    html_3 = web('https://fssp.gov.ru/mvv_fssp/')
    get_content(html_3)
    del html_3
    html_4 = web('https://pfr.gov.ru/info/af/')
    get_content(html_4)
    del html_4
    dest_mail = input("Введите почту получателя: ")
    dest_mail = str(dest_mail)
    print("Производится отправка сообщения...")
    send_email(smev3_news + next_smev3_news, dest_mail)
    """
    input("Нажмите Enter для выхода")


newparser()
