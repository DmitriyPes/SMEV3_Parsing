# -*- coding: cp1251 -*-
import os
from selenium import webdriver
import time
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
import traceback
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import ElementClickInterceptedException
import pandas
import numpy as np

"""
from selenium.webdriver import ActionChains

actions = ActionChains(driver)
actions.move_to_element(element).click().perform()
"""

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
chromedriver = ROOT_DIR + r"\chromedriver"

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
        self.options.add_argument('headless')
        self.options.add_experimental_option("prefs", self.preferences)
        self.options.add_experimental_option('excludeSwitches', ['enable-logging'])

    def set_preferencies(self):
        self.preferences = {"download.default_directory": str(ROOT_DIR) + r"\downloads", "safebrowsing.enabled": "false"}
    def close_connection(self):
        self.browser.close()

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
    item = re.sub(r'<a class="collapsed" data-parent="#accordion" data-toggle="collapse" >', '<span style="color:#FF0000" class="collapsed" data-parent="#accordion" data-toggle="collapse" >', str(item))
    item = re.sub(r'</a></h2>', '</span></h2>', str(item))
    return item

def get_themes(html):
    i = 3
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
    return themes

def check_file(file_name):
    for i in os.listdir(str(ROOT_DIR) + r"\downloads"):
        if file_name in i:
            return True
        continue
    return False

def open_tab(html, link):
    try:
        html.browser.execute_script("window.open();")
        main_window = html.browser.current_window_handle
        html.browser.switch_to.window(html.browser.window_handles[1])
        html.browser.get(link)
        time.sleep(0.1)
    except:
        print(traceback.format_exc())
    name = get_newsname(html)
    html.browser.close()
    html.browser.switch_to.window(main_window)
    return name

def get_newsname(html):
    name = html.browser.find_element_by_xpath('//*[@id="infotable"]/table/tbody/tr[1]').text
    name = re.sub(r'Наименование ', '', str(name))
    name = '<h3> ' + 'Новость содержит информацию о следующем ВС - ' + name + ' </h3>'
    return name

def check_vs(vs_info, name):
    for code,vs_name in vs_info.items():
        if name == vs_name:
            return code
        return ""

def get_content(html):
    news = []
    if html.link == 'https://smev3.gosuslugi.ru/portal':
        soup = BeautifulSoup(html.get_html_text(),"lxml")
        items = soup.find_all("div", {"class": "info-section gray-container"})
        for item in items:
            item_tmp = item
            color_check1 = item_tmp.find("span", style = 'color:#FF0000')
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
                news.append(tmps)
        return news
    if html.link == 'https://smev3.gosuslugi.ru/portal/news.jsp':
        path = './Виды сведений.xlsx'
        vs_info = excel_connect(path)
        i = 2
        while True:
            xpath = '/html/body/div/div/div/div/div[1]/div/div[1]/div[2]/ul/li[i]/a'
            xpath = xpath.replace('[i]', '[' + str(i) + ']')
            try:
                html.browser.find_element_by_xpath(xpath).click()
                time.sleep(0.1)
                soup = BeautifulSoup(html.get_html_text(), "lxml")
                items = soup.find_all("div", {"class": "panel panel-news"})
                for item in items:
                    links_with_text = []
                    for a in item.find_all('a', href=True):
                        if a.text:
                            if str(a['href']).startswith('https://smev3.gosuslugi.ru') == True:
                                links_with_text.append(a['href'])
                    if links_with_text != []:
                        name = open_tab(html, links_with_text[0])
                    else: name = ''
                    vs_add = check_vs(vs_info, name)
                    if vs_add != '':
                        name = name + "Данные найдены в таблице видов сведений со следующим кодом - " + vs_add
                    else: name = name + "Данные не найдены в таблице видов сведений"
                    if html.filter == True:
                        if item.get_text().find(html.filter_word) != -1:
                            news.append(name + ' ' + edit_news(item))
                    else:
                        news.append(name + ' ' + edit_news(item))
            except (NoSuchElementException, ElementClickInterceptedException):
                break
            i+=1
            time.sleep(0.1)
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
    email = 'tesmail.test@yandex.ru'
    password = 'testmail1'

    server = smtplib.SMTP('smtp.yandex.ru', 587)
    server.set_debuglevel(False)
    server.ehlo()
    server.starttls()
    server.login(email, password)

    from_addr = 'tesmail.test@yandex.ru'
    msg_body = '<html> <body>'
    for i in range (0, len(news)):
        msg_body = str(msg_body) + ''.join(map(str,news[i]))
    msg_body = str(msg_body) + '</body></html>'

    """
    msg_body_new = ''
    msg_body_new = msg_body_new.join(map(str,msg_body))
    print(str(msg_body_new))
    """

    msg = MIMEText(str(msg_body), 'html', 'utf-8')
    msg['Subject'] = 'Это письмо от Песоцкого Дмитрия'
    msg['From'] = from_addr
    msg['To'] = toaddr

    server.sendmail(email, toaddr, msg.as_string())
    print("Сообщение успешно отправлено")
    server.quit()

def check_theme(theme_name, themes):
    for theme in themes:
        if theme_name == theme:
            return True
    return False


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
    #for new in smev3_news:
       #print(new)
    del html_1
    print("Новости отобраны.")
    dest_mail = input("Введите почту получателя: ")
    dest_mail = str(dest_mail)
    print("Производится отправка сообщения...")
    send_email(smev3_news + next_smev3_news, dest_mail)
    input("Нажмите Enter для выхода")
newparser()